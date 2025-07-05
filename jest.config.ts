// jest.config.js
/** @type {import('ts-jest').JestConfigWithTsJest} */
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'jsdom', // Use jsdom for browser-like environment (needed for FileReader)
  testMatch: ["**/__tests__/**/*.test.ts"], // Specify where your test files are
  moduleFileExtensions: ['ts', 'js', 'json', 'node'],
  // Setup for mocking global objects like FileReader
  setupFilesAfterEnv: ["<rootDir>/jest.setup.js"],
};