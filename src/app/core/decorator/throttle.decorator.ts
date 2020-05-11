export function throttle(threshhold: number = 10): MethodDecorator {
  return function (
    target: any,
    propertyKey: string,
    descriptor: PropertyDescriptor
  ) {
    const timeoutKey = Symbol(`__timeout__${propertyKey}`);
    const original = descriptor.value;
    descriptor.value = function (...args: any[]) {
      if (!this[timeoutKey]) {
        this[timeoutKey] = setTimeout(() => {
          original.apply(this, args);
          clearTimeout(this[timeoutKey]);
          this[timeoutKey] = null;
        }, threshhold);
      }
    };
    return descriptor;
  };
}
