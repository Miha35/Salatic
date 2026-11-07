import { defineConfig } from 'vitest/config';

/** Accuracy evaluations with real models and data (slow; see tests/eval). */
export default defineConfig({
  test: {
    include: ['tests/eval/**/*.eval.test.ts'],
    environment: 'node',
    testTimeout: 7_200_000,
    hookTimeout: 1_800_000,
    fileParallelism: false,
  },
});
