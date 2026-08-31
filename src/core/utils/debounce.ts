export interface DebouncedFunction<T extends (...args: any[]) => any> {
  (...args: Parameters<T>): void;
  cancel(): void;
  flush(): void;
}

export function debounce<T extends (...args: any[]) => any>(
  fn: T,
  waitMs: number
): DebouncedFunction<T> {
  let timer: number | null = null;
  let lastArgs: Parameters<T> | null = null;

  const debounced = (...args: Parameters<T>) => {
    lastArgs = args;
    if (timer !== null) {
      window.clearTimeout(timer);
    }
    timer = window.setTimeout(() => {
      timer = null;
      if (lastArgs) {
        fn(...lastArgs);
        lastArgs = null;
      }
    }, waitMs);
  };

  debounced.cancel = () => {
    if (timer !== null) {
      window.clearTimeout(timer);
      timer = null;
    }
    lastArgs = null;
  };

  debounced.flush = () => {
    if (timer !== null) {
      window.clearTimeout(timer);
      timer = null;
      if (lastArgs) {
        fn(...lastArgs);
        lastArgs = null;
      }
    }
  };

  return debounced;
}
