/**
 * Runtime configuration.
 * For a real deployment, prefer injecting these via build-time env vars.
 */
export const API_BASE_URL = (window as any).__API_BASE_URL__ || window.location.origin;
