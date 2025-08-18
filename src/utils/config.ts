/**
 * Configuration management for Excel add-in
 * Handles environment-based settings for Django backend and Firebase
 */

export interface AppConfig {
  djangoBackendUrl: string;
  firebaseConfig: FirebaseConfig;
  isDevelopment: boolean;
  isProduction: boolean;
}

export interface FirebaseConfig {
  apiKey: string;
  authDomain: string;
  projectId: string;
  storageBucket: string;
  messagingSenderId: string;
  appId: string;
}

// Environment detection
const isDevelopment = () => {
  // Check various indicators for development environment
  return (
    window.location.hostname === 'localhost' ||
    window.location.hostname === '127.0.0.1' ||
    window.location.port === '3000' ||
    window.location.protocol === 'http:' ||
    process.env.NODE_ENV === 'development'
  );
};

// Firebase configuration
const getFirebaseConfig = (): FirebaseConfig => {
  // In a real environment, these would come from process.env or build-time configuration
  // For Excel add-in, we'll provide them directly from your configuration
  return {
    apiKey: "AIzaSyDIByC_lBZSGOZKM0uYBrfjMkqUVVKpnK8",
    authDomain: "jir-automator.firebaseapp.com",
    projectId: "jir-automator",
    storageBucket: "jir-automator.firebasestorage.app",
    messagingSenderId: "959393365413",
    appId: "1:959393365413:web:00a46fe380873c9a84bc4c"
  };
};

// Django backend URL configuration
const getDjangoBackendUrl = (): string => {
  if (isDevelopment()) {
    return 'http://localhost:8000';
  } else {
    return 'https://github-jirventures-hyperion-server-doealo2leq-uc.a.run.app';
  }
};

// Main configuration object
export const getAppConfig = (): AppConfig => {
  const isDevEnv = isDevelopment();
  
  return {
    djangoBackendUrl: getDjangoBackendUrl(),
    firebaseConfig: getFirebaseConfig(),
    isDevelopment: isDevEnv,
    isProduction: !isDevEnv,
  };
};

// Configuration getters for easy access
export const getDjangoUrl = (): string => getAppConfig().djangoBackendUrl;
export const getFirebaseApiKey = (): string => getAppConfig().firebaseConfig.apiKey;
export const getFirebaseAuthDomain = (): string => getAppConfig().firebaseConfig.authDomain;
export const getFirebaseProjectId = (): string => getAppConfig().firebaseConfig.projectId;

// Environment helpers
export const isDevEnvironment = (): boolean => getAppConfig().isDevelopment;
export const isProdEnvironment = (): boolean => getAppConfig().isProduction;

