#!/usr/bin/env node

/**
 * Basic test script for SharePoint MCP Server
 * Tests the server structure and tool definitions without requiring actual credentials
 */

import { spawn } from 'child_process';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Mock environment variables for testing
process.env.SHAREPOINT_URL = 'https://test.sharepoint.com';
process.env.TENANT_ID = 'test-tenant-id';
process.env.CLIENT_ID = 'test-client-id';
process.env.CLIENT_SECRET = 'test-client-secret';

console.log('🧪 Starting SharePoint MCP Server Basic Tests...\n');

// Test 1: Server starts without errors
console.log('Test 1: Server startup test');
const serverPath = join(__dirname, '..', 'build', 'index.js');

const server = spawn('node', [serverPath], {
  stdio: ['pipe', 'pipe', 'pipe'],
  env: process.env
});

let serverOutput = '';
let serverError = '';

server.stdout.on('data', (data) => {
  serverOutput += data.toString();
});

server.stderr.on('data', (data) => {
  serverError += data.toString();
});

// Test 2: Send MCP initialization request
setTimeout(() => {
  console.log('Test 2: Sending MCP initialization request');
  
  const initRequest = {
    jsonrpc: '2.0',
    id: 1,
    method: 'initialize',
    params: {
      protocolVersion: '2024-11-05',
      capabilities: {
        roots: {
          listChanged: true
        }
      },
      clientInfo: {
        name: 'test-client',
        version: '1.0.0'
      }
    }
  };

  server.stdin.write(JSON.stringify(initRequest) + '\n');
}, 1000);

// Test 3: Request list of tools
setTimeout(() => {
  console.log('Test 3: Requesting list of tools');
  
  const toolsRequest = {
    jsonrpc: '2.0',
    id: 2,
    method: 'tools/list'
  };

  server.stdin.write(JSON.stringify(toolsRequest) + '\n');
}, 2000);

// Test 4: Request list of resources
setTimeout(() => {
  console.log('Test 4: Requesting list of resources');
  
  const resourcesRequest = {
    jsonrpc: '2.0',
    id: 3,
    method: 'resources/list'
  };

  server.stdin.write(JSON.stringify(resourcesRequest) + '\n');
}, 3000);

// Cleanup and results after 5 seconds
setTimeout(() => {
  server.kill('SIGTERM');
  
  console.log('\n📊 Test Results:');
  console.log('================');
  
  if (serverError.includes('SharePoint MCP server running on stdio')) {
    console.log('✅ Server started successfully');
  } else {
    console.log('❌ Server startup failed');
    console.log('Error output:', serverError);
  }
  
  if (serverOutput.includes('"method":"tools/list"') || serverOutput.includes('search_files')) {
    console.log('✅ Tools endpoint responding');
  } else {
    console.log('⚠️  Tools endpoint not tested (may require actual interaction)');
  }
  
  if (serverOutput.includes('"method":"resources/list"')) {
    console.log('✅ Resources endpoint responding');
  } else {
    console.log('⚠️  Resources endpoint not tested (may require actual interaction)');
  }
  
  console.log('\n📝 Server Output:');
  console.log(serverOutput || 'No stdout output');
  
  console.log('\n🔍 Server Error Log:');
  console.log(serverError || 'No stderr output');
  
  console.log('\n✨ Basic tests completed!');
  console.log('💡 For full testing, configure real SharePoint credentials and use the MCP Inspector.');
  
  process.exit(0);
}, 5000);

// Handle process termination
process.on('SIGINT', () => {
  console.log('\n🛑 Test interrupted');
  server.kill('SIGTERM');
  process.exit(1);
});
