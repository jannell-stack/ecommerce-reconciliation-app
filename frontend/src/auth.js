const STORAGE_KEY = 'app_password';

export function isAuthenticated() {
  return Boolean(sessionStorage.getItem(STORAGE_KEY));
}

export function getStoredPassword() {
  return sessionStorage.getItem(STORAGE_KEY) || '';
}

export function setStoredPassword(password) {
  sessionStorage.setItem(STORAGE_KEY, password);
}

export function clearAuth() {
  sessionStorage.removeItem(STORAGE_KEY);
}
