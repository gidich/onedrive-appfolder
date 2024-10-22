/** @type {import('ts-jest').JestConfigWithTsJest} **/
module.exports = {
  roots: ["<rootDir>/."],
  preset: 'ts-jest',
  globalSetup: "<rootDir>/src/test/global-setup.ts",
  reporters: [
    "default",
    ["./node_modules/jest-html-reporter", {
      "pageTitle": "Owner Community - Data Access Report"
    }]
  ],

  testEnvironment: "node",
  testMatch: ['**/**/*.test.ts'],
  coverageReporters: ['json', 'lcov'],


  transform: {
    "^.+.tsx?$": ["ts-jest",{}],
  },
  collectCoverageFrom: [
      "**/*.{js,jsx,ts}",
      "!**/node_modules/**",
      "!**/dist/**"
    ]




};