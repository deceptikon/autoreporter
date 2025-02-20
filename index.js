const axios = require('axios');
const ExcelJS = require('exceljs');

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

// Function to copy merged cells
function copyMergedCells(sourceSheet, targetSheet, rowIndex, offset) {
  Object.keys(sourceSheet._merges).forEach(range => {
    const merge = sourceSheet._merges[range];
    const targetTopLeftRow = merge.tl.row + rowIndex - offset;
    const targetTopLeftCol = merge.tl.col;
    const targetBottomRightRow = merge.br.row + rowIndex - offset;
    const targetBottomRightCol = merge.br.col;

    // Check if the target range is already merged
    let isAlreadyMerged = false;
    Object.keys(targetSheet._merges).forEach(existingRange => {
      const existingMerge = targetSheet._merges[existingRange];
      if (existingMerge.tl.row === targetTopLeftRow && existingMerge.tl.col === targetTopLeftCol &&
        existingMerge.br.row === targetBottomRightRow && existingMerge.br.col === targetBottomRightCol) {
        isAlreadyMerged = true;
      }
    });

    if (!isAlreadyMerged) {
      // targetSheet.mergeCells(targetTopLeftRow, targetTopLeftCol, targetBottomRightRow, targetBottomRightCol);
    }
  });
}

async function generateReport(data) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Report');

  // Load the sample workbook
  const sampleWorkbook = new ExcelJS.Workbook();
  await sampleWorkbook.xlsx.readFile('sample.xlsx');
  const sampleSheet = sampleWorkbook.getWorksheet(1);

  // Force calculation
  sampleWorkbook.calcProperties.fullCalcOnLoad = true;
  await sampleWorkbook.xlsx.readFile('sample.xlsx');

  // Log merged cells from the sample Excel file
  console.log(`Sample Sheet Merges:`, sampleSheet.merges);

  const groupedData = data.reduce((acc, item) => {
    const prefix = item.userName.split('_')[0];
    if (!acc[prefix]) acc[prefix] = [];
    acc[prefix].push(item);
    return acc;
  }, {});

  let rowIndex = 1;

  for (const prefix in groupedData) {
    // Copy the first two rows from sample.xlsx
    for (let i = 1; i <= 2; i++) {
      const sourceRow = sampleSheet.getRow(i);
      const targetRow = worksheet.getRow(rowIndex);

      // Copy all cells from source row to target row
      sourceRow.eachCell((cell, colNumber) => {
        targetRow.getCell(colNumber).value = cell.value;
        targetRow.getCell(colNumber).style = JSON.parse(JSON.stringify(cell.style)); // Deep copy of style
      });

      // Copy merged cells
      copyMergedCells(sampleSheet, worksheet, rowIndex, i);

      rowIndex++;
    }

    const group = groupedData[prefix].sort((a, b) => a.userName.localeCompare(b.userName));
    worksheet.getCell(`A${rowIndex}`).value = prefix;
    rowIndex++;

    group.forEach((item, index) => {
      const fullName = item.fullName.replace(/^.*?_/, '');
      worksheet.getCell(`A${rowIndex}`).value = index + 1;
      worksheet.getCell(`B${rowIndex}`).value = fullName;
      worksheet.getCell(`C${rowIndex}`).value = item.data.month;
      rowIndex++;
    });

    const totalTraffic = group.reduce((sum, item) => sum + item.data.month, 0);
    worksheet.getCell(`B${rowIndex}`).value = 'Total';
    worksheet.getCell(`C${rowIndex}`).value = totalTraffic;
    rowIndex++;

    // Copy rows 4 to 7 from sample.xlsx
    for (let i = 4; i <= 7; i++) {
      const sourceRow = sampleSheet.getRow(i);
      const targetRow = worksheet.getRow(rowIndex);

      // Copy all cells from source row to target row
      sourceRow.eachCell((cell, colNumber) => {
        targetRow.getCell(colNumber).value = cell.value;
        targetRow.getCell(colNumber).style = JSON.parse(JSON.stringify(cell.style)); // Deep copy of style
      });

      // Copy merged cells
      copyMergedCells(sampleSheet, worksheet, rowIndex, i);

      rowIndex++;
    }

    rowIndex++;
  }

  await workbook.xlsx.writeFile('report.xlsx');
}

// Function to create Excel tables for each group
function createExcelTablesForGroups(groups) {
  Object.keys(groups).forEach(groupName => {
    const data = groups[groupName];
    // createXLSFile(data, `${groupName}_traffic_usage.xlsx`);
  });
}

// Example usage
(async () => {
  const baseURL = `https://${process.argv[4] || '192.168.0.101'}:4081/admin/api/jsonrpc/`;
  const username = process.argv[2];
  const password = process.argv[3];
  if (!baseURL || !username || !password) {
    console.error('Error: Please provide baseURL, username, and password as command line arguments.');
    process.exit(1);
  }
  const api = createApiInstance(baseURL);
  await login(api, username, password);
  const response = await getTrafficUsage(api);

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
    await generateReport(users); // Call the function with your data
  } else {
    console.error('Error: No data returned from the API.');
  }
})();

module.exports = {
  login,
  queryKerio
};
