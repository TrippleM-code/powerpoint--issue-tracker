export type Result<T> =
  | { ok: true; value: T }
  | { ok: false; error: Error };

export function ok<T>(value: T): Result<T> {
  return { ok: true, value };
}

export function fail(error: unknown): Result<never> {
  return {
    ok: false,
    error: error instanceof Error ? error : new Error(String(error)),
  };
}
