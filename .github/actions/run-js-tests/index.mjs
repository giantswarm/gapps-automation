import {getInput, setFailed} from '@actions/core';

async function runTest() {
    try {
        // many tests require an implementation of fetch() which is only in stable node v18
        if (typeof fetch !== 'function') {
            globalThis.fetch = (await import('node-fetch')).default;
        }

        // `who-to-greet` input defined in action metadata file
        const test_name = getInput('test_name', {required: true});

        await import(`../../../tests/${test_name}`);
    } catch (error) {
        setFailed(error.message);
    }
}

runTest();
