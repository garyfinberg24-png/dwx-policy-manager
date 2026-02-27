/**
 * Jest Configuration for Policy Manager SPFx Project
 *
 * Uses ts-jest to compile TypeScript test files.
 * Mocks SCSS modules and SPFx/PnP framework imports.
 */
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'jsdom',
  roots: ['<rootDir>/src'],
  testMatch: [
    '**/__tests__/**/*.test.ts',
    '**/__tests__/**/*.test.tsx',
    '**/*.test.ts',
    '**/*.test.tsx',
  ],
  transform: {
    '^.+\\.tsx?$': [
      'ts-jest',
      {
        tsconfig: {
          target: 'es2017',
          module: 'commonjs',
          moduleResolution: 'node',
          jsx: 'react',
          esModuleInterop: true,
          skipLibCheck: true,
          resolveJsonModule: true,
          lib: ['es2015', 'es2017', 'dom'],
          types: ['jest'],
        },
      },
    ],
  },
  moduleNameMapper: {
    // Mock SCSS/CSS modules
    '\\.module\\.s?css$': '<rootDir>/src/__mocks__/styleMock.js',
    '\\.s?css$': '<rootDir>/src/__mocks__/styleMock.js',

    // Mock SPFx framework modules
    '^@microsoft/sp-(.*)$': '<rootDir>/src/__mocks__/spfxMock.js',

    // Mock PnP modules
    '^@pnp/(.*)$': '<rootDir>/src/__mocks__/pnpMock.js',

    // Mock @dwx/core
    '^@dwx/core$': '<rootDir>/src/__mocks__/dwxCoreMock.js',
  },
  setupFiles: ['<rootDir>/src/__mocks__/setupGlobals.js'],
  // Ignore node_modules except for specific packages if needed
  transformIgnorePatterns: ['/node_modules/'],
  // Coverage configuration
  collectCoverageFrom: [
    'src/services/**/*.ts',
    '!src/services/**/*.test.ts',
    '!src/**/__mocks__/**',
  ],
};
