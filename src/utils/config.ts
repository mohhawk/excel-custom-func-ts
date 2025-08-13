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

// Configuration validation
export const validateConfig = (): { isValid: boolean; errors: string[] } => {
  const config = getAppConfig();
  const errors: string[] = [];

  // Validate Django URL
  if (!config.djangoBackendUrl) {
    errors.push('Django backend URL is not configured');
  } else {
    try {
      new URL(config.djangoBackendUrl);
    } catch {
      errors.push('Django backend URL is not a valid URL');
    }
  }

  // Validate Firebase configuration
  const requiredFirebaseFields = ['apiKey', 'authDomain', 'projectId', 'storageBucket', 'messagingSenderId', 'appId'];
  for (const field of requiredFirebaseFields) {
    if (!config.firebaseConfig[field as keyof FirebaseConfig]) {
      errors.push(`Firebase ${field} is not configured`);
    }
  }

  return {
    isValid: errors.length === 0,
    errors,
  };
};

// Debug helper
export const getConfigDebugInfo = (): string => {
  const config = getAppConfig();
  const validation = validateConfig();
  
  return JSON.stringify({
    environment: config.isDevelopment ? 'development' : 'production',
    djangoUrl: config.djangoBackendUrl,
    firebaseProjectId: config.firebaseConfig.projectId,
    validation: validation,
    hostname: window.location.hostname,
    port: window.location.port,
    protocol: window.location.protocol,
  }, null, 2);
};

// Storage keys for user overrides
const CONFIG_OVERRIDE_KEY = 'excel_addin_config_override';

// Allow user configuration override (for testing/development)
export interface ConfigOverride {
  djangoBackendUrl?: string;
  firebaseApiKey?: string;
}

export const saveConfigOverride = (override: ConfigOverride): void => {
  try {
    localStorage.setItem(CONFIG_OVERRIDE_KEY, JSON.stringify(override));
  } catch (error) {
    console.warn('Failed to save config override:', error);
  }
};

export const getConfigOverride = (): ConfigOverride | null => {
  try {
    const stored = localStorage.getItem(CONFIG_OVERRIDE_KEY);
    return stored ? JSON.parse(stored) : null;
  } catch (error) {
    console.warn('Failed to load config override:', error);
    return null;
  }
};

export const clearConfigOverride = (): void => {
  try {
    localStorage.removeItem(CONFIG_OVERRIDE_KEY);
  } catch (error) {
    console.warn('Failed to clear config override:', error);
  }
};

// Get effective configuration (with user overrides)
export const getEffectiveConfig = (): AppConfig => {
  const baseConfig = getAppConfig();
  const override = getConfigOverride();
  
  if (!override) {
    return baseConfig;
  }
  
  return {
    ...baseConfig,
    djangoBackendUrl: override.djangoBackendUrl || baseConfig.djangoBackendUrl,
    firebaseConfig: {
      ...baseConfig.firebaseConfig,
      apiKey: override.firebaseApiKey || baseConfig.firebaseConfig.apiKey,
    },
  };
};
