#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

// Read the current version.json
const versionPath = path.join(__dirname, '..', 'version.json');
const versionData = JSON.parse(fs.readFileSync(versionPath, 'utf8'));

// Parse current version
const [major, minor, patch] = versionData.version.split('.').map(Number);

// Increment patch version
const newPatch = patch + 1;
const newVersion = `${major}.${minor}.${newPatch}`;

// Store old version for logging
const oldVersion = versionData.version;

// Update version.json
versionData.version = newVersion;
versionData.buildDate = new Date().toISOString().split('T')[0];

// Write back to file
fs.writeFileSync(versionPath, JSON.stringify(versionData, null, 2));

console.log(`✅ Version bumped from ${oldVersion} to ${newVersion}`);
console.log(`📅 Build date updated to ${versionData.buildDate}`);
console.log(`📝 Don't forget to update the description in version.json if needed!`);
