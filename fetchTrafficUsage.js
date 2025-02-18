const axios = require('axios');
const fs = require('fs');

// Function to create an Axios instance
function createApiInstance(baseURL) {
  return axios.create({
    baseURL: baseURL,
    headers: {
      'Content-Type': 'application/json'
    },
    httpsAgent: new (require('https').Agent)({ rejectUnauthorized: false }) // Ignore SSL certificate errors
  });
}

let token = null;
let sessionCookie = null;

// Function to login and get the token
async function login(api, username, password) {
  try {
    const response = await api.post('', {
      "jsonrpc": "2.0",
      "id": 1,
      "method": "Session.login",
      "params": {
        "userName": username,
        "password": password,
        "application": {
          "name": "Example App",
          "vendor": "Kerio Technologies s.r.o.",
          "version": "1.0"
        }
      }
    });
    token = response.data.result.token;
    sessionCookie = response.headers['set-cookie'].find(cookie => cookie.startsWith('SESSION_CONTROL_WEBADMIN'));
    api.defaults.headers.common['X-Token'] = token;
    api.defaults.headers.common['Cookie'] = sessionCookie;
    console.log('Login successful, token:', token);
  } catch (error) {
    console.error('Error logging in:', error);
  }
}

// Function to query Kerio API
async function queryKerio(api, method, params = {}) {
  if (!token) {
    console.error('Error: Not authenticated. Please login first.');
    return;
  }
  try {
    console.log('Request headers:', api.defaults.headers.common);
    const response = await api.post('', {
      "jsonrpc": "2.0",
      "id": 1,
      "method": method,
      "params": params
    });
    console.log(response.data);
    return response;
  } catch (error) {
    console.error('Error querying Kerio API:', error);
  }
}

// Function to query traffic usage statistics
async function getTrafficUsage(api) {
  return await queryKerio(api, "UserStatistics.get", {
    query: {
      start: 0,
      limit: -1,
      orderBy: [{ columnName: "userName", direction: "Asc" }],
    },
    refresh: true // Set refresh to true
  });
}

// Main function to login, query traffic usage, and save results to a JSON file
async function fetchTrafficUsage(baseURL, username, password) {
  const api = createApiInstance(baseURL);
  await login(api, username, password);
  const response = await getTrafficUsage(api);

  // Assuming response.data.result.list contains the traffic usage data
  if (response && response.data && response.data.result && response.data.result.list) {
    const users = response.data.result.list;
    if (!users || users.length === 0) {
      console.error('Error: No users returned.');
      return;
    }
    fs.writeFileSync('traffic_usage.json', JSON.stringify(users, null, 2));
    console.log('Traffic usage data saved to traffic_usage.json');
  } else {
    console.error('Error: No data returned from the API.');
  }
}

// Example usage
const baseURL = `https://${process.argv[4] || '192.168.0.101'}:4081/admin/api/jsonrpc/`;
const username = process.argv[2];
const password = process.argv[3];
if (!baseURL || !username || !password) {
  console.error('Error: Please provide baseURL, username, and password as command line arguments.');
  process.exit(1);
}
fetchTrafficUsage(baseURL, username, password);
