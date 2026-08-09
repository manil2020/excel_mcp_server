import { defineConfig } from 'vitest/config';

export default defineConfig({
  // Under vitest, HyperFormula's CommonJS bundle hits a circular-dep bug during
  // default-language registration. Force resolution to the ES bundle for tests.
  resolve: {
    alias: {
      hyperformula: 'hyperformula/es/index.js',
    },
  },
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'node',
    testTimeout: 20_000,
    server: {
      deps: {
        inline: [/^hyperformula/],
      },
    },
    coverage: {
      provider: 'v8',
      reporter: ['text', 'html'],
      include: ['src/**/*.ts'],
    },
  },
});
