import assert from 'node:assert/strict';

const {createServer} = await import('http');
const {PersonioClientV1, PersonioAuthV1} = (await import('../lib-output/lib.js')).default;

const mockServer = createServer((request, response) => {
    console.log("request: ", (request.method || '').toUpperCase(), request.url);
    let status = 200;
    let content = '';
    if (request.url.includes('/auth')) {
        content = JSON.stringify({success: true, data: {token: '123bla'}});
    } else if (request.url.includes('/company/employee')) {
        if (request.headers?.authorization?.includes('123bla')) {
            content = JSON.stringify({success: true, data: [{type: 'Employee', attributes: {'id': {value: 'xyz'}}}]});
        } else {
            status = 401;
        }
    } else {
        status = 404;
    }
    response.writeHead(status, {'Content-Type': 'application/json'});
    response.end(content, 'utf-8');
});

try {
    const baseUrl = await new Promise(resolve => mockServer.listen({
        host: '127.0.0.1',
        port: 0
    }, () => resolve('http://127.0.0.1:' + mockServer.address().port)));

    const personioAuth = new PersonioAuthV1("blablaClient", "bluSecret", baseUrl);
    const personio = new PersonioClientV1(personioAuth, baseUrl);

    const employees = await personio.getPersonioJson('/company/employee');

    assert(Array.isArray(employees), 'result is not an array');
    assert(employees.find(employee => employee?.attributes?.id?.value === 'xyz'), 'expected employee not found in result');
} finally {
    mockServer.close();
}
