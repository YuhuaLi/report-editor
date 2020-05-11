export function debounce(delay: number = 300): MethodDecorator {
  return (target: any, propertyKey: string, descriptor: PropertyDescriptor) => {
    const timeoutKey = Symbol(`__timeout__${propertyKey}`);
    const original = descriptor.value;
    descriptor.value = (...args: any[]) => {
      clearTimeout(this[timeoutKey]);
      this[timeoutKey] = setTimeout(() => original.apply(this, args), delay);
    };
    return descriptor;
  };
}
