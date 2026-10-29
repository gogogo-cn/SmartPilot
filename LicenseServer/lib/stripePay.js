'use strict';

const paymentConfig = require('./paymentConfig');

function priceOk(id) {
    return typeof id === 'string' && /^price_[A-Za-z0-9]+$/.test(id) && !/^price_x+$/i.test(id);
}

function ready() {
    const key = paymentConfig.get(paymentConfig.SECRET_KEY) || '';
    if (!/^sk_(live|test)_/.test(key)) return false;
    return priceOk(paymentConfig.get(paymentConfig.PRICE_MONTHLY)) && priceOk(paymentConfig.get(paymentConfig.PRICE_YEARLY));
}

function getStripe() {
    const key = paymentConfig.get(paymentConfig.SECRET_KEY);
    if (!key) return null;
    // Lazy require so local tests work without node_modules until npm install
    const Stripe = require('stripe');
    return new Stripe(key);
}

async function createCheckoutSession({ user, plan, successUrl, cancelUrl }) {
    if (!ready()) {
        const err = new Error('Stripe checkout is not available yet');
        err.code = 'checkout_not_ready';
        throw err;
    }
    const stripe = getStripe();
    const price =
        plan === 'pro_yearly'
            ? paymentConfig.get(paymentConfig.PRICE_YEARLY)
            : paymentConfig.get(paymentConfig.PRICE_MONTHLY);

    const session = await stripe.checkout.sessions.create({
        mode: 'subscription',
        customer_email: user.email || undefined,
        line_items: [{ price, quantity: 1 }],
        success_url: successUrl || paymentConfig.get('STRIPE_SUCCESS_URL'),
        cancel_url: cancelUrl || paymentConfig.get('STRIPE_CANCEL_URL'),
        client_reference_id: user.id,
        metadata: {
            userId: user.id,
            plan
        },
        subscription_data: {
            metadata: {
                userId: user.id,
                plan
            }
        }
    });
    return { id: session.id, url: session.url };
}

async function createPortalSession({ customerId, returnUrl }) {
    const stripe = getStripe();
    if (!stripe) {
        const err = new Error('STRIPE_SECRET_KEY not configured');
        err.code = 'stripe_not_configured';
        throw err;
    }
    if (!customerId) {
        const err = new Error('No Stripe customer on file');
        err.code = 'stripe_no_customer';
        throw err;
    }
    const session = await stripe.billingPortal.sessions.create({
        customer: customerId,
        return_url: returnUrl || paymentConfig.get('STRIPE_SUCCESS_URL')
    });
    return { url: session.url };
}

function constructWebhookEvent(rawBody, signature) {
    const stripe = getStripe();
    if (!stripe) {
        const err = new Error('STRIPE_SECRET_KEY not configured');
        err.code = 'stripe_not_configured';
        throw err;
    }
    const secret = paymentConfig.get('STRIPE_WEBHOOK_SECRET');
    if (!secret) {
        const err = new Error('STRIPE_WEBHOOK_SECRET not configured');
        err.code = 'stripe_webhook_secret_missing';
        throw err;
    }
    return stripe.webhooks.constructEvent(rawBody, signature, secret);
}

module.exports = {
    ready,
    getStripe,
    createCheckoutSession,
    createPortalSession,
    constructWebhookEvent
};
