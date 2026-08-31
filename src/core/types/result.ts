/**
 * Represents the outcome of an operation that can succeed or fail.
 */
export type Result<T, E = Error> =
  | { success: true; data: T }
  | { success: false; error: E };

export const Result = {
  ok<T>(data: T): Result<T, never> {
    return { success: true, data };
  },

  fail<E>(error: E): Result<never, E> {
    return { success: false, error };
  },

  fromThrowable<T>(fn: () => T): Result<T, Error> {
    try {
      return Result.ok(fn());
    } catch (err) {
      return Result.fail(err instanceof Error ? err : new Error(String(err)));
    }
  },

  async fromAsyncThrowable<T>(fn: () => Promise<T>): Promise<Result<T, Error>> {
    try {
      const data = await fn();
      return Result.ok(data);
    } catch (err) {
      return Result.fail(err instanceof Error ? err : new Error(String(err)));
    }
  },
};
