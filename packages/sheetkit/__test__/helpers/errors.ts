import { isResourceLimitError, type ResourceLimitError } from '../../index.js';

/**
 * Runs `action` and returns the `Error` it throws.
 * Fails when `action` returns normally or throws a value that is not an `Error`.
 */
export function captureError(action: () => void): Error {
  try {
    action();
  } catch (error) {
    if (error instanceof Error) return error;

    throw new Error('Expected action to throw an Error instance', { cause: error });
  }

  throw new Error('Expected action to throw');
}

/**
 * Awaits `pending` and returns the `Error` it rejects with.
 * Fails when `pending` resolves or rejects with a value that is not an `Error`.
 */
export async function captureRejection<T>(pending: Promise<T>): Promise<Error> {
  try {
    await pending;
  } catch (error) {
    if (error instanceof Error) return error;

    throw new Error('Expected promise to reject with an Error instance', { cause: error });
  }

  throw new Error('Expected promise to reject');
}

/**
 * Runs `action` and returns the `ResourceLimitError` it throws.
 * Any other thrown value is rethrown unchanged so the test reports it.
 */
export function captureResourceLimitError(action: () => void): ResourceLimitError {
  try {
    action();
  } catch (error) {
    if (isResourceLimitError(error)) return error;

    throw error;
  }

  throw new Error('Expected a resource limit error');
}
