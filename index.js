const axios = require('axios');
const xlsx = require('xlsx'); // Add this line

// Base URL for Kerio API
const baseURL = 'https://192.168.0.101:4081/admin/api/jsonrpc/';

// Create an Axios instance
const api = axios.create({
  baseURL: baseURL,
  headers: {
    'Content-Type': 'application/json'
  }
});

let token = null;
let sessionCookie = null;

// Function to login and get the token
async function login(username, password) {
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
async function queryKerio(method, params = {}) {
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

// Function to create an XLS file
function createXLSFile(data, fileName) {
  const worksheet = xlsx.utils.json_to_sheet(data);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  xlsx.writeFile(workbook, fileName);
}

// Function to query traffic usage statistics
async function getTrafficUsage() {
  return await queryKerio("UserStatistics.get", {
    query: {
      start: 0,
      limit: -1,
      orderBy: [{ columnName: "userName", direction: "Asc" }],
    },
    refresh: true // Set refresh to true
  });
}

// Function to split users into groups based on name prefixes
function splitUsersIntoGroups(users) {
  const groups = {};
  users.forEach(user => {
    if (user.type !== 'UserStatisticUser') return; // Filter out non-user statistics
    const prefix = user.userName.split('_')[0]; // Assuming prefix is before an underscore
    if (!groups[prefix]) {
      groups[prefix] = [];
    }
    groups[prefix].push(user);
  });
  return groups;
}

// Function to create Excel tables for each group
function createExcelTablesForGroups(groups) {
  Object.keys(groups).forEach(groupName => {
    const data = groups[groupName];
    createXLSFile(data, `${groupName}_traffic_usage.xlsx`);
  });
}

// Example usage
(async () => {
  await login('admin', '1q'); // Replace with your actual username and password
  const response = await getTrafficUsage();

  // Assuming response.data.result.list contains the traffic usage data
  if (response && response.data && response.data.result && response.data.result.list) {
    const users = response.data.result.list;
    console.warn(users);
    if (!users || users.length === 0) {
      console.error('Error: No users returned.');
      return;
    }
    const groups = splitUsersIntoGroups(users);
    createExcelTablesForGroups(groups);
  } else {
    console.error('Error: No data returned from the API.');
  }
})();

module.exports = {
  login,
  queryKerio
};
