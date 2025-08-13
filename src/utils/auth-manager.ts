/**
 * Authentication manager for Firebase + Django backend integration
 */

// ========== Interfaces ==========

export interface AuthCredentials {
  email: string;
  password: string;
}

export interface RegisterCredentials {
  email: string;
  password: string;
  firstName: string;
  lastName: string;
}

export interface FirebaseUser {
  uid: string;
  email: string;
  displayName: string;
  emailVerified: boolean;
  photoURL?: string;
}

export interface AuthToken {
  idToken: string;
  refreshToken: string;
  expirationTime: number;
  localId: string;
}

export interface UserInfo {
  id: number;
  username: string;
  email: string;
  first_name: string;
  last_name: string;
  credit_balance: number;
  subscription_tier: string;
  is_active: boolean;
  last_login: string;
}

export interface OLAPConnection {
  id: number;
  name: string;
  olap_type: 'hyperion' | 'ssas' | 'tm1' | 'jedox';
  server_url: string;
  application?: string;
  is_active: boolean;
  created_at: string;
  last_used?: string;
  description?: string;
}

export interface AuthState {
  isAuthenticated: boolean;
  user: UserInfo | null;
  token: AuthToken | null;
  connections: OLAPConnection[];
  lastActivity: Date | null;
}

export interface LoginResult {
  success: boolean;
  message: string;
  firebaseUser?: FirebaseUser;
  user?: UserInfo;
  token?: AuthToken;
  connections?: OLAPConnection[];
}

export interface RegisterResult {
  success: boolean;
  message: string;
  firebaseUser?: FirebaseUser;
  needsEmailVerification?: boolean;
}

export interface ApiResponse<T = any> {
  success: boolean;
  data?: T;
  error?: string;
  message?: string;
  status?: number;
}

// ========== Authentication Manager Class ==========

export class AuthManager {
  private baseUrl: string;
  private authState: AuthState;
  private refreshTimer: number | null = null;
  private readonly FIREBASE_API_KEY = 'your-firebase-api-key'; // Will be configurable
  private readonly FIREBASE_AUTH_URL = 'https://identitytoolkit.googleapis.com/v1/accounts';
  private readonly TOKEN_STORAGE_KEY = 'firebase_auth_token';
  private readonly USER_STORAGE_KEY = 'firebase_user_info';
  private readonly DJANGO_USER_STORAGE_KEY = 'django_user_info';
  private readonly CONNECTIONS_STORAGE_KEY = 'olap_connections';

  constructor(baseUrl: string = 'http://localhost:8000', firebaseApiKey?: string) {
    this.baseUrl = baseUrl.replace(/\/$/, ''); // Remove trailing slash
    if (firebaseApiKey) {
      (this as any).FIREBASE_API_KEY = firebaseApiKey;
    }
    
    this.authState = {
      isAuthenticated: false,
      user: null,
      token: null,
      connections: [],
      lastActivity: null,
    };

    // Try to restore session from storage
    this.restoreSession();
  }

  /**
   * Configure Firebase API key
   */
  setFirebaseConfig(apiKey: string): void {
    (this as any).FIREBASE_API_KEY = apiKey;
  }

  /**
   * Login with email and password through Firebase
   */
  async login(credentials: AuthCredentials): Promise<LoginResult> {
    try {
      // Step 1: Authenticate with Firebase
      const firebaseResponse = await fetch(`${this.FIREBASE_AUTH_URL}:signInWithPassword?key=${this.FIREBASE_API_KEY}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          email: credentials.email,
          password: credentials.password,
          returnSecureToken: true,
        }),
      });

      const firebaseData = await firebaseResponse.json();

      if (!firebaseResponse.ok) {
        return {
          success: false,
          message: this.getFirebaseErrorMessage(firebaseData.error?.message) || 'Login failed',
        };
      }

      // Step 2: Store Firebase token and user info
      const authToken: AuthToken = {
        idToken: firebaseData.idToken,
        refreshToken: firebaseData.refreshToken,
        expirationTime: Date.now() + (parseInt(firebaseData.expiresIn) * 1000),
        localId: firebaseData.localId,
      };

      const firebaseUser: FirebaseUser = {
        uid: firebaseData.localId,
        email: firebaseData.email,
        displayName: firebaseData.displayName || '',
        emailVerified: firebaseData.emailVerified || false,
        photoURL: firebaseData.photoUrl,
      };

      this.authState.token = authToken;
      this.authState.isAuthenticated = true;
      this.authState.lastActivity = new Date();

      // Step 3: Get or create Django user profile
      const djangoUserResult = await this.syncWithDjangoBackend(firebaseUser, authToken.idToken);
      
      if (djangoUserResult.success) {
        this.authState.user = djangoUserResult.data.user;
        this.authState.connections = djangoUserResult.data.connections || [];
      }

      // Persist to storage
      this.saveToStorage();
      this.saveFirebaseUser(firebaseUser);

      // Set up token refresh
      this.setupTokenRefresh();

      return {
        success: true,
        message: 'Login successful',
        firebaseUser: firebaseUser,
        user: this.authState.user,
        token: authToken,
        connections: this.authState.connections,
      };

    } catch (error) {
      return {
        success: false,
        message: `Network error: ${error.message}`,
      };
    }
  }

  /**
   * Register new user with Firebase
   */
  async register(credentials: RegisterCredentials): Promise<RegisterResult> {
    try {
      // Step 1: Create user with Firebase
      const firebaseResponse = await fetch(`${this.FIREBASE_AUTH_URL}:signUp?key=${this.FIREBASE_API_KEY}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          email: credentials.email,
          password: credentials.password,
          returnSecureToken: true,
        }),
      });

      const firebaseData = await firebaseResponse.json();

      if (!firebaseResponse.ok) {
        return {
          success: false,
          message: this.getFirebaseErrorMessage(firebaseData.error?.message) || 'Registration failed',
        };
      }

      // Step 2: Update user profile with display name
      const displayName = `${credentials.firstName} ${credentials.lastName}`;
      await this.updateFirebaseProfile(firebaseData.idToken, {
        displayName: displayName,
      });

      // Step 3: Send email verification
      await this.sendEmailVerification(firebaseData.idToken);

      const firebaseUser: FirebaseUser = {
        uid: firebaseData.localId,
        email: firebaseData.email,
        displayName: displayName,
        emailVerified: false,
        photoURL: undefined,
      };

      // Step 4: Create Django user profile
      const authToken: AuthToken = {
        idToken: firebaseData.idToken,
        refreshToken: firebaseData.refreshToken,
        expirationTime: Date.now() + (parseInt(firebaseData.expiresIn) * 1000),
        localId: firebaseData.localId,
      };

      await this.syncWithDjangoBackend(firebaseUser, authToken.idToken);

      return {
        success: true,
        message: 'Registration successful! Please check your email for verification.',
        firebaseUser: firebaseUser,
        needsEmailVerification: true,
      };

    } catch (error) {
      return {
        success: false,
        message: `Network error: ${error.message}`,
      };
    }
  }

  /**
   * Logout and clear session
   */
  async logout(): Promise<void> {
    try {
      // Clear local state (Firebase logout is implicit)
      this.clearSession();
    } catch (error) {
      console.warn('Logout failed:', error.message);
    }
  }

  /**
   * Sync user data with Django backend
   */
  private async syncWithDjangoBackend(firebaseUser: FirebaseUser, idToken: string): Promise<ApiResponse> {
    try {
      const response = await fetch(`${this.baseUrl}/api/auth/firebase-sync/`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${idToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          firebase_uid: firebaseUser.uid,
          email: firebaseUser.email,
          display_name: firebaseUser.displayName,
          email_verified: firebaseUser.emailVerified,
        }),
      });

      const data = await response.json();

      if (!response.ok) {
        return {
          success: false,
          error: data.error || 'Failed to sync with backend',
        };
      }

      return {
        success: true,
        data: data,
      };

    } catch (error) {
      return {
        success: false,
        error: `Backend sync failed: ${error.message}`,
      };
    }
  }

  /**
   * Update Firebase user profile
   */
  private async updateFirebaseProfile(idToken: string, profileData: { displayName?: string; photoUrl?: string }): Promise<void> {
    await fetch(`${this.FIREBASE_AUTH_URL}:update?key=${this.FIREBASE_API_KEY}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        idToken: idToken,
        displayName: profileData.displayName,
        photoUrl: profileData.photoUrl,
        returnSecureToken: false,
      }),
    });
  }

  /**
   * Send email verification
   */
  private async sendEmailVerification(idToken: string): Promise<void> {
    await fetch(`${this.FIREBASE_AUTH_URL}:sendOobCode?key=${this.FIREBASE_API_KEY}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        requestType: 'VERIFY_EMAIL',
        idToken: idToken,
      }),
    });
  }

  /**
   * Get Firebase error message
   */
  private getFirebaseErrorMessage(errorCode: string): string {
    const errorMessages: { [key: string]: string } = {
      'EMAIL_EXISTS': 'An account with this email already exists.',
      'EMAIL_NOT_FOUND': 'No account found with this email address.',
      'INVALID_PASSWORD': 'Invalid password.',
      'USER_DISABLED': 'This account has been disabled.',
      'WEAK_PASSWORD': 'Password is too weak. Please choose a stronger password.',
      'INVALID_EMAIL': 'Invalid email address.',
      'OPERATION_NOT_ALLOWED': 'This operation is not allowed.',
      'TOO_MANY_ATTEMPTS_TRY_LATER': 'Too many failed attempts. Please try again later.',
    };

    return errorMessages[errorCode] || 'Authentication failed. Please try again.';
  }

  /**
   * Save Firebase user to storage
   */
  private saveFirebaseUser(firebaseUser: FirebaseUser): void {
    try {
      localStorage.setItem(this.USER_STORAGE_KEY, JSON.stringify(firebaseUser));
    } catch (error) {
      console.warn('Failed to save Firebase user to storage:', error.message);
    }
  }

  /**
   * Get stored Firebase user
   */
  getFirebaseUser(): FirebaseUser | null {
    try {
      const userData = localStorage.getItem(this.USER_STORAGE_KEY);
      return userData ? JSON.parse(userData) : null;
    } catch (error) {
      console.warn('Failed to retrieve Firebase user from storage:', error.message);
      return null;
    }
  }

  /**
   * Refresh Firebase ID token
   */
  async refreshToken(): Promise<boolean> {
    if (!this.authState.token?.refreshToken) {
      return false;
    }

    try {
      const response = await fetch(`https://securetoken.googleapis.com/v1/token?key=${this.FIREBASE_API_KEY}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          grant_type: 'refresh_token',
          refresh_token: this.authState.token.refreshToken,
        }),
      });

      if (!response.ok) {
        // Refresh token is invalid, need to re-login
        this.clearSession();
        return false;
      }

      const data = await response.json();
      
      // Update token with new ID token
      this.authState.token = {
        idToken: data.id_token,
        refreshToken: data.refresh_token,
        expirationTime: Date.now() + (parseInt(data.expires_in) * 1000),
        localId: data.user_id,
      };

      this.authState.lastActivity = new Date();
      this.saveToStorage();
      this.setupTokenRefresh();

      return true;

    } catch (error) {
      console.error('Token refresh failed:', error.message);
      this.clearSession();
      return false;
    }
  }

  /**
   * Get current user info with updated credit balance
   */
  async getUserInfo(): Promise<UserInfo | null> {
    if (!this.authState.isAuthenticated) {
      return null;
    }

    try {
      const response = await this.makeAuthenticatedRequest('/api/user/profile/');
      
      if (response.success && response.data) {
        this.authState.user = response.data;
        this.saveToStorage();
        return response.data;
      }

      return this.authState.user;

    } catch (error) {
      console.error('Failed to fetch user info:', error.message);
      return this.authState.user;
    }
  }

  /**
   * Get OLAP connections for the current user
   */
  async getConnections(): Promise<OLAPConnection[]> {
    if (!this.authState.isAuthenticated) {
      return [];
    }

    try {
      const response = await this.makeAuthenticatedRequest('/api/olap/connections/');
      
      if (response.success && response.data) {
        this.authState.connections = response.data;
        this.saveToStorage();
        return response.data;
      }

      return this.authState.connections;

    } catch (error) {
      console.error('Failed to fetch connections:', error.message);
      return this.authState.connections;
    }
  }

  /**
   * Create a new OLAP connection
   */
  async createConnection(connectionData: {
    name: string;
    olap_type: OLAPConnection['olap_type'];
    server_url: string;
    application?: string;
    username: string;
    password: string;
    description?: string;
  }): Promise<ApiResponse<OLAPConnection>> {
    if (!this.authState.isAuthenticated) {
      return { success: false, error: 'Not authenticated' };
    }

    try {
      const response = await this.makeAuthenticatedRequest('/api/olap/connections/', {
        method: 'POST',
        body: JSON.stringify(connectionData),
      });

      if (response.success) {
        // Refresh connections list
        await this.getConnections();
      }

      return response;

    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  /**
   * Delete an OLAP connection
   */
  async deleteConnection(connectionId: number): Promise<ApiResponse> {
    if (!this.authState.isAuthenticated) {
      return { success: false, error: 'Not authenticated' };
    }

    try {
      const response = await this.makeAuthenticatedRequest(`/api/olap/connections/${connectionId}/`, {
        method: 'DELETE',
      });

      if (response.success) {
        // Remove from local state
        this.authState.connections = this.authState.connections.filter(c => c.id !== connectionId);
        this.saveToStorage();
      }

      return response;

    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  /**
   * Make an authenticated API request using Firebase ID token
   */
  async makeAuthenticatedRequest(
    endpoint: string, 
    options: RequestInit = {}
  ): Promise<ApiResponse> {
    if (!this.authState.isAuthenticated || !this.authState.token) {
      return { success: false, error: 'Not authenticated' };
    }

    // Check if token needs refresh
    if (this.isTokenExpiring()) {
      const refreshed = await this.refreshToken();
      if (!refreshed) {
        return { success: false, error: 'Authentication expired' };
      }
    }

    try {
      const response = await fetch(`${this.baseUrl}${endpoint}`, {
        ...options,
        headers: {
          'Authorization': `Bearer ${this.authState.token.idToken}`,
          'Content-Type': 'application/json',
          ...options.headers,
        },
      });

      // Handle authentication errors
      if (response.status === 401) {
        // Try to refresh token once
        const refreshed = await this.refreshToken();
        if (refreshed) {
          // Retry request with new token
          return this.makeAuthenticatedRequest(endpoint, options);
        } else {
          this.clearSession();
          return { success: false, error: 'Authentication expired' };
        }
      }

      const data = await response.json();

      if (!response.ok) {
        return {
          success: false,
          error: data.error || data.message || 'Request failed',
          status: response.status,
        };
      }

      this.authState.lastActivity = new Date();

      return {
        success: true,
        data: data,
        status: response.status,
      };

    } catch (error) {
      return { success: false, error: `Network error: ${error.message}` };
    }
  }

  /**
   * Check current authentication state
   */
  getAuthState(): AuthState {
    return { ...this.authState };
  }

  /**
   * Check if user is authenticated
   */
  isAuthenticated(): boolean {
    return this.authState.isAuthenticated && this.authState.token !== null;
  }

  /**
   * Get current user
   */
  getCurrentUser(): UserInfo | null {
    return this.authState.user;
  }

  /**
   * Get available connections
   */
  getAvailableConnections(): OLAPConnection[] {
    return [...this.authState.connections];
  }

  /**
   * Get connection by ID
   */
  getConnection(id: number): OLAPConnection | null {
    return this.authState.connections.find(c => c.id === id) || null;
  }

  /**
   * Update base URL
   */
  updateBaseUrl(url: string): void {
    this.baseUrl = url.replace(/\/$/, '');
  }

  /**
   * Check if token is expiring soon (within 5 minutes)
   */
  private isTokenExpiring(): boolean {
    if (!this.authState.token) {
      return true;
    }

    const now = Date.now();
    const expirationTime = this.authState.token.expirationTime;
    const timeRemaining = expirationTime - now;

    return timeRemaining < 5 * 60 * 1000; // 5 minutes
  }

  /**
   * Set up automatic token refresh
   */
  private setupTokenRefresh(): void {
    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer);
    }

    if (!this.authState.token) {
      return;
    }

    // Refresh token 5 minutes before it expires
    const now = Date.now();
    const expirationTime = this.authState.token.expirationTime;
    const refreshIn = Math.max(0, expirationTime - now - (5 * 60 * 1000)); // 5 minutes before expiry
    
    this.refreshTimer = window.setTimeout(async () => {
      await this.refreshToken();
    }, refreshIn);
  }

  /**
   * Save authentication state to local storage
   */
  private saveToStorage(): void {
    try {
      if (this.authState.token) {
        localStorage.setItem(this.TOKEN_STORAGE_KEY, JSON.stringify(this.authState.token));
      }
      
      if (this.authState.user) {
        localStorage.setItem(this.DJANGO_USER_STORAGE_KEY, JSON.stringify(this.authState.user));
      }
      
      if (this.authState.connections.length > 0) {
        localStorage.setItem(this.CONNECTIONS_STORAGE_KEY, JSON.stringify(this.authState.connections));
      }
    } catch (error) {
      console.warn('Failed to save authentication state to storage:', error.message);
    }
  }

  /**
   * Restore authentication state from local storage
   */
  private restoreSession(): void {
    try {
      const tokenData = localStorage.getItem(this.TOKEN_STORAGE_KEY);
      const firebaseUserData = localStorage.getItem(this.USER_STORAGE_KEY);
      const djangoUserData = localStorage.getItem(this.DJANGO_USER_STORAGE_KEY);
      const connectionsData = localStorage.getItem(this.CONNECTIONS_STORAGE_KEY);

      if (tokenData && firebaseUserData) {
        const token = JSON.parse(tokenData);
        
        // Check if token is still valid
        if (token.expirationTime > Date.now()) {
          this.authState.token = token;
          this.authState.user = djangoUserData ? JSON.parse(djangoUserData) : null;
          this.authState.connections = connectionsData ? JSON.parse(connectionsData) : [];
          this.authState.isAuthenticated = true;
          this.authState.lastActivity = new Date();

          // Set up token refresh
          this.setupTokenRefresh();
        } else {
          // Token expired, clear session
          this.clearSession();
        }
      }
    } catch (error) {
      console.warn('Failed to restore session from storage:', error.message);
      this.clearSession();
    }
  }

  /**
   * Clear authentication state and storage
   */
  private clearSession(): void {
    this.authState = {
      isAuthenticated: false,
      user: null,
      token: null,
      connections: [],
      lastActivity: null,
    };

    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer);
      this.refreshTimer = null;
    }

    try {
      localStorage.removeItem(this.TOKEN_STORAGE_KEY);
      localStorage.removeItem(this.USER_STORAGE_KEY);
      localStorage.removeItem(this.DJANGO_USER_STORAGE_KEY);
      localStorage.removeItem(this.CONNECTIONS_STORAGE_KEY);
    } catch (error) {
      console.warn('Failed to clear storage:', error.message);
    }
  }
}

// ========== Utility Functions ==========

/**
 * Create authentication manager instance
 */
export function createAuthManager(baseUrl?: string, firebaseApiKey?: string): AuthManager {
  return new AuthManager(baseUrl, firebaseApiKey);
}

/**
 * Global authentication manager instance
 */
let globalAuthManager: AuthManager | null = null;

/**
 * Get global authentication manager instance
 */
export function getAuthManager(baseUrl?: string, firebaseApiKey?: string): AuthManager {
  if (!globalAuthManager) {
    globalAuthManager = new AuthManager(baseUrl, firebaseApiKey);
  } else {
    if (baseUrl) {
      globalAuthManager.updateBaseUrl(baseUrl);
    }
    if (firebaseApiKey) {
      globalAuthManager.setFirebaseConfig(firebaseApiKey);
    }
  }
  
  return globalAuthManager;
}

/**
 * Clear global authentication manager (for testing)
 */
export function clearGlobalAuthManager(): void {
  globalAuthManager = null;
}
