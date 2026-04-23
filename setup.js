#!/usr/bin/env node
/**
 * G-Stack Asset Command - Client Setup Script
 * Creates a standalone Google Apps Script project for fleet/asset tracking
 *
 * Usage: node setup.js
 *
 * True North Data Strategies
 * jacob@truenorthstrategyops.com
 */

const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { execSync } = require('child_process');

// Module Configuration
const MODULE_CONFIG = {
  name: 'G-Stack Asset Command',
  description: 'Fleet & Asset Tracking System',
  projectSuffix: 'G-Stack-Asset-Command',
  projectType: 'sheets',
  defaultIcon: '🚗',

  // Files to include in deployment
  templateFiles: [
    'Code.gs',
    'FunctionRunner.gs',
    'Dashboard.html',
    'appsscript.json'
  ],

  // File rename mappings
  fileRenames: {},

  // Legacy string replacements
  legacyReplacements: {
    'ACME Equipment Rentals': 'companyName',
    'demo@example.com': 'companyEmail'
  },

  // Next steps
  nextSteps: [
    'Open the Google Sheet that was created',
    'Refresh the page to see the 🚗 Tools menu',
    'Run "1. Build Complete Template" from the menu',
    'Run "2. Setup Dashboard" from the menu',
    'Run "3. Add Test Data" (optional)',
    'Go to Config sheet and update your settings'
  ]
};

// =========================================
// Readline Utilities
// =========================================

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function ask(question) {
  return new Promise(resolve => rl.question(question, resolve));
}

// =========================================
// Main Setup Function
// =========================================

async function main() {
  console.log('\n========================================');
  console.log(`   ${MODULE_CONFIG.name} - Client Setup`);
  console.log(`   ${MODULE_CONFIG.description}`);
  console.log('========================================\n');
  console.log('This wizard will create a customized Google Apps Script project');
  console.log('with your company branding and deploy it to Google Workspace.\n');

  // Gather client information
  const answers = await collectCompanyInfo();

  // Display summary
  displaySummary(answers);

  // Confirm
  const confirm = await ask('\nProceed with setup? (y/n): ');
  if (confirm.toLowerCase() !== 'y') {
    console.log('Setup cancelled.');
    rl.close();
    process.exit(0);
  }

  // Create output directory
  const safeName = createSafeName(answers.companyName);
  const outputDir = await createOutputDirectory(safeName);

  // Process template files
  console.log('\nProcessing template files...');
  const replacements = buildReplacementMap(answers, safeName);
  processFiles(outputDir, replacements);

  // Create config.gs with company settings
  createConfigFile(outputDir, answers);

  // Create .clasp.json.template
  createClaspTemplate(outputDir);

  // Close readline before clasp operations
  rl.close();

  // Deploy with clasp
  const projectTitle = `${answers.shortName} ${MODULE_CONFIG.projectSuffix}`;
  const claspJson = deployWithClasp(outputDir, projectTitle);

  // Display completion message
  displayComplete(claspJson, answers);
}

// =========================================
// Company Info Collection
// =========================================

async function collectCompanyInfo() {
  const answers = {};

  // Required: Company Name
  answers.companyName = await ask('Business Name (e.g., "ABC Rentals"): ');
  if (!answers.companyName.trim()) {
    console.log('Error: Business name is required.');
    rl.close();
    process.exit(1);
  }
  answers.companyName = answers.companyName.trim();

  // Optional: Short Name
  const shortNameInput = await ask(`Short Name for menus [${answers.companyName}]: `);
  answers.shortName = shortNameInput.trim() || answers.companyName;

  // Optional: Menu Icon
  const iconInput = await ask(`Menu Icon emoji [${MODULE_CONFIG.defaultIcon}]: `);
  answers.menuIcon = iconInput.trim() || MODULE_CONFIG.defaultIcon;

  // Optional: Email for alerts
  answers.companyEmail = (await ask('Email for alerts (optional): ')).trim();

  // Optional: Phone
  answers.companyPhone = (await ask('Company Phone (optional): ')).trim();

  return answers;
}

// =========================================
// Display Functions
// =========================================

function displaySummary(answers) {
  console.log('\n--- Configuration Summary ---');
  console.log(`Module: ${MODULE_CONFIG.name}`);
  console.log(`Business Name: ${answers.companyName}`);
  console.log(`Short Name: ${answers.shortName}`);
  console.log(`Menu Icon: ${answers.menuIcon}`);
  console.log(`Alert Email: ${answers.companyEmail || '(not set)'}`);
  console.log(`Phone: ${answers.companyPhone || '(not set)'}`);
}

function displayComplete(claspJson, answers) {
  console.log('\n========================================');
  console.log('           SETUP COMPLETE!');
  console.log('========================================\n');
  console.log(`${MODULE_CONFIG.name} created successfully!\n`);

  if (claspJson && claspJson.scriptId) {
    console.log('Script Editor:');
    console.log(`  https://script.google.com/d/${claspJson.scriptId}/edit\n`);
  }

  console.log('Next Steps:');
  MODULE_CONFIG.nextSteps.forEach((step, i) => {
    console.log(`  ${i + 1}. ${step}`);
  });
  console.log('');
  console.log('Support: jacob@truenorthstrategyops.com\n');
}

// =========================================
// File Processing
// =========================================

function createSafeName(companyName) {
  return companyName
    .replace(/[^a-zA-Z0-9]/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
}

async function createOutputDirectory(safeName) {
  const dirName = `${safeName}-${MODULE_CONFIG.projectSuffix}`;
  const outputDir = path.join(__dirname, '..', '..', dirName);

  if (fs.existsSync(outputDir)) {
    const overwrite = await ask(`\nDirectory "${dirName}" exists. Overwrite? (y/n): `);
    if (overwrite.toLowerCase() !== 'y') {
      console.log('Setup cancelled.');
      rl.close();
      process.exit(0);
    }
    fs.rmSync(outputDir, { recursive: true });
  }

  console.log(`\nCreating project in: ${outputDir}`);
  fs.mkdirSync(outputDir, { recursive: true });

  return outputDir;
}

function buildReplacementMap(answers, safeName) {
  const replacements = {
    // Standard placeholders
    '{{COMPANY_NAME}}': answers.companyName,
    '{{SHORT_NAME}}': answers.shortName,
    '{{MENU_ICON}}': answers.menuIcon,
    '{{COMPANY_EMAIL}}': answers.companyEmail || 'alerts@example.com',
    '{{COMPANY_PHONE}}': answers.companyPhone || '',
    '{{SAFE_NAME}}': safeName,
    '{{TRUENORTH_EMAIL}}': 'jacob@truenorthstrategyops.com'
  };

  // Add legacy replacements
  for (const [search, answerKey] of Object.entries(MODULE_CONFIG.legacyReplacements)) {
    if (answerKey === 'companyName') {
      replacements[search] = answers.companyName;
    } else if (answerKey === 'companyEmail') {
      replacements[search] = answers.companyEmail || 'alerts@example.com';
    }
  }

  return replacements;
}

function processFiles(outputDir, replacements) {
  for (const filename of MODULE_CONFIG.templateFiles) {
    const srcPath = path.join(__dirname, filename);
    let destFilename = MODULE_CONFIG.fileRenames[filename] || filename;

    for (const [search, replace] of Object.entries(replacements)) {
      destFilename = destFilename.split(search).join(replace);
    }

    const destPath = path.join(outputDir, destFilename);

    if (fs.existsSync(srcPath)) {
      let content = fs.readFileSync(srcPath, 'utf8');

      for (const [search, replace] of Object.entries(replacements)) {
        content = content.split(search).join(replace);
      }

      fs.writeFileSync(destPath, content);
      console.log(`  ✓ ${destFilename}`);
    } else {
      console.log(`  ⚠ ${filename} not found, skipping`);
    }
  }
}

function createConfigFile(outputDir, answers) {
  const configContent = `/**
 * Company Configuration - ${answers.companyName}
 * Auto-generated by G-Stack Asset Command Setup
 */

const COMPANY_CONFIG = {
  NAME: '${answers.companyName}',
  SHORT_NAME: '${answers.shortName}',
  MENU_ICON: '${answers.menuIcon}',
  ALERT_EMAIL: '${answers.companyEmail || ""}',
  PHONE: '${answers.companyPhone || ""}',
  TRUENORTH_EMAIL: 'jacob@truenorthstrategyops.com'
};

function getCompanyName() {
  return COMPANY_CONFIG.NAME;
}

function getAlertEmail() {
  return COMPANY_CONFIG.ALERT_EMAIL;
}
`;

  fs.writeFileSync(path.join(outputDir, 'config.gs'), configContent);
  console.log('  ✓ config.gs (generated)');
}

function createClaspTemplate(outputDir) {
  const template = {
    scriptId: 'WILL_BE_SET_BY_CLASP_CREATE',
    rootDir: './'
  };
  const templatePath = path.join(outputDir, '.clasp.json.template');
  fs.writeFileSync(templatePath, JSON.stringify(template, null, 2));
  console.log('  ✓ .clasp.json.template');
}

// =========================================
// Clasp Deployment
// =========================================

function deployWithClasp(outputDir, projectTitle) {
  console.log('\n--- Creating Google Apps Script Project ---');

  try {
    process.chdir(outputDir);

    try {
      execSync('clasp --version', { stdio: 'pipe' });
    } catch (e) {
      console.log('\nNote: clasp CLI not found or not logged in.');
      console.log('Install with: npm install -g @google/clasp');
      console.log('Login with: clasp login');
      console.log('\nFiles have been prepared in the output directory.');
      console.log('You can manually run these commands later:');
      console.log(`  cd "${outputDir}"`);
      console.log(`  clasp create --title "${projectTitle}" --type ${MODULE_CONFIG.projectType}`);
      console.log('  clasp push --force');
      return null;
    }

    console.log('Creating new spreadsheet and script...');
    execSync(`clasp create --title "${projectTitle}" --type ${MODULE_CONFIG.projectType}`, { stdio: 'inherit' });

    console.log('\nPushing files to Google Apps Script...');
    execSync('clasp push --force', { stdio: 'inherit' });

    const claspJsonPath = path.join(outputDir, '.clasp.json');
    if (fs.existsSync(claspJsonPath)) {
      return JSON.parse(fs.readFileSync(claspJsonPath, 'utf8'));
    }

    return null;

  } catch (error) {
    console.error('\nError during clasp operations:', error.message);
    console.log('\nFiles have been prepared. You can manually run:');
    console.log(`  cd "${outputDir}"`);
    console.log(`  clasp create --title "${projectTitle}" --type ${MODULE_CONFIG.projectType}`);
    console.log('  clasp push --force');
    return null;
  }
}

// =========================================
// Run
// =========================================

main().catch(err => {
  console.error('Setup failed:', err);
  rl.close();
  process.exit(1);
});

