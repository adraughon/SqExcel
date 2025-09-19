#!/usr/bin/env node

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

// Read the current version
const versionPath = path.join(__dirname, '..', 'version.json');
const versionData = JSON.parse(fs.readFileSync(versionPath, 'utf8'));
const version = versionData.version;

// Check if a summary argument was provided
const summaryArg = process.argv[2];

if (summaryArg) {
  // Summary provided as argument
  const commitMessage = `Release v${version}: ${summaryArg}`;
  commitChanges(commitMessage);
} else {
  // No summary provided, prompt for input
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  rl.question(`📝 Enter release summary for v${version}: `, (summary) => {
    rl.close();
    
    if (!summary || summary.trim() === '') {
      console.error('❌ Release summary is required. Aborting release.');
      process.exit(1);
    }
    
    const commitMessage = `Release v${version}: ${summary.trim()}`;
    commitChanges(commitMessage);
  });
}

function commitChanges(commitMessage) {
  try {
    execSync(`git commit -m "${commitMessage}"`, { stdio: 'inherit' });
    console.log(`✅ Committed with message: "${commitMessage}"`);
    
    // Push to remote
    console.log('🚀 Pushing to remote repository...');
    execSync('git push', { stdio: 'inherit' });
    console.log('✅ Successfully pushed to remote repository');
  } catch (error) {
    console.error('❌ Failed to commit or push:', error.message);
    process.exit(1);
  }
}
