export function createI18n(getString: (key: string) => string) {
  return (key: string, values: Record<string, string | number> = {}) =>
    Object.entries(values).reduce(
      (text, [name, value]) => text.split(`{${name}}`).join(String(value)),
      getString(key),
    );
}
