/** Numéro de version de Grillon (celui de package.json au moment de la construction). */
export const appVersion = (): string => (typeof __APP_VERSION__ === 'string' ? __APP_VERSION__ : '');
