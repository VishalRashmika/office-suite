/**
 * Represents a resource that can be disposed/cleaned up.
 */
export interface IDisposable {
  dispose(): void;
}

/**
 * Manages a collection of disposable objects.
 */
export class CompositeDisposable implements IDisposable {
  private disposables = new Set<IDisposable>();
  private isDisposed = false;

  add(...disposables: (IDisposable | (() => void))[]): void {
    if (this.isDisposed) {
      for (const d of disposables) {
        if (typeof d === "function") d();
        else d.dispose();
      }
      return;
    }

    for (const d of disposables) {
      if (typeof d === "function") {
        this.disposables.add({ dispose: d });
      } else {
        this.disposables.add(d);
      }
    }
  }

  dispose(): void {
    if (this.isDisposed) return;
    this.isDisposed = true;
    for (const d of this.disposables) {
      try {
        d.dispose();
      } catch (err) {
        console.error("Error disposing resource:", err);
      }
    }
    this.disposables.clear();
  }
}
