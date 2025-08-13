/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import { 
  DataPreprocessor, 
  createDataProcessor, 
  estimateProcessingCost,
  PreprocessedData,
  ChunkingStrategy 
} from '../utils/data-processor';
import { 
  ResponseParser, 
  createResponseParser, 
  ChunkResponse,
  MappingResult 
} from '../utils/response-parser';
import { 
  AuthManager, 
  getAuthManager, 
  UserInfo, 
  OLAPConnection,
  LoginResult,
  RegisterResult,
  AuthCredentials,
  RegisterCredentials,
  FirebaseUser
} from '../utils/auth-manager';
import {
  getAppConfig,
  isDevEnvironment,
  validateConfig,
  getConfigDebugInfo,
  saveConfigOverride,
  getConfigOverride,
  clearConfigOverride,
  getEffectiveConfig
} from '../utils/config';

interface EPMSettings {
  serverUrl: string;
  application: string;
  username: string;
  password: string;
}

interface OperationProgress {
  isActive: boolean;
  operationId: string;
  totalChunks: number;
  completedChunks: number;
  currentChunk: number;
  estimatedCells: number;
  estimatedCost: number;
  startTime: Date;
  canCancel: boolean;
}

// Global instances
let authManager: AuthManager;
let dataProcessor: DataPreprocessor;
let responseParser: ResponseParser;
let currentOperation: OperationProgress | null = null;

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  
  // Initialize components
  initializeComponents();
  
  // Load saved settings and check authentication
  loadSavedSettings();
  checkAuthenticationState();
  
  // Add event listeners
  setupEventListeners();
});

/**
 * Initialize all components
 */
function initializeComponents() {
  // Validate configuration first
  const configValidation = validateConfig();
  if (!configValidation.isValid) {
    console.error('Configuration validation failed:', configValidation.errors);
    showNotification(`Configuration error: ${configValidation.errors.join(', ')}`, true);
  }

  // Get effective configuration (environment + user overrides)
  const config = getEffectiveConfig();
  
  // Log environment info for debugging
  if (isDevEnvironment()) {
    console.log('Excel Add-in Configuration:', getConfigDebugInfo());
  }
  
  // Initialize authentication manager with environment-based configuration
  authManager = getAuthManager(config.djangoBackendUrl, config.firebaseConfig.apiKey);
  
  // Initialize data processor with default chunking strategy
  dataProcessor = createDataProcessor({
    maxPayloadSize: 1024 * 1024, // 1MB
    maxCellsPerChunk: 10000,
    chunkByDimension: true,
    preserveStructure: true,
  });
  
  // Initialize response parser
  responseParser = createResponseParser({
    validateIntegrity: true,
    handleMissingChunks: 'error',
    sortByChunkIndex: true,
    mergeStrategy: 'smart',
  });
  
  // Load saved configuration and populate UI
  loadSavedConfiguration();
  
  // Update UI with current configuration
  updateConfigurationUI();
}

/**
 * Setup all event listeners
 */
function setupEventListeners() {
  // Authentication events
  document.getElementById("login-btn").onclick = handleLogin;
  document.getElementById("register-btn").onclick = handleRegister;
  document.getElementById("logout-btn").onclick = handleLogout;
  document.getElementById("refresh-user-info").onclick = handleRefreshUserInfo;
  
  // Form switching events
  document.getElementById("show-register-btn").onclick = showRegisterForm;
  document.getElementById("show-login-btn").onclick = showLoginForm;
  
  // Connection management events
  document.getElementById("save-connection").onclick = saveConnection;
  
  // Operation events
  document.getElementById("refresh-data").onclick = handleRefreshData;
  document.getElementById("estimate-cost").onclick = handleEstimateCost;
  document.getElementById("cancel-operation").onclick = handleCancelOperation;
  
  // Settings events
  document.getElementById("max-cells-chunk").onchange = updateChunkingStrategy;
  document.getElementById("max-payload-size").onchange = updateChunkingStrategy;
  document.getElementById("chunk-by-dimension").onchange = updateChunkingStrategy;

  // Removed config change events for non-existent fields
}
/**
* Check current authentication state and update UI
*/
async function checkAuthenticationState() {
if (authManager.isAuthenticated()) {
  const firebaseUser = authManager.getFirebaseUser();
  const djangoUser = authManager.getCurrentUser();
  
  if (firebaseUser) {
    showAuthenticatedState(firebaseUser, djangoUser);
    await loadConnections();
  } else {
    showUnauthenticatedState();
  }
  } else {
    showUnauthenticatedState();
  }
}

/**
 * Update configuration UI with current settings
 */
function updateConfigurationUI() {
  const config = getEffectiveConfig();
  const override = getConfigOverride();
  
  // Removed setting values for non-existent fields:
  // (document.getElementById('django-url') as HTMLInputElement).value = config.djangoBackendUrl;
  // (document.getElementById('firebase-api-key') as HTMLInputElement).value = config.firebaseConfig.apiKey;
  
  // Show environment indicator
  const envIndicator = document.createElement('div');
  envIndicator.className = 'ms-font-s environment-indicator';

  
  if (isDevEnvironment()) {
    envIndicator.classList.add('env-development');
    envIndicator.innerHTML = '🔧 Development Environment - Using localhost backend';
  } else {
    envIndicator.classList.add('env-production');
    envIndicator.innerHTML = '🚀 Production Environment - Using cloud backend';
  }
  
  // Add override indicator if present
  if (override) {
    const overrideIndicator = document.createElement('div');
    overrideIndicator.className = 'ms-font-xs config-override';
    overrideIndicator.innerHTML = '⚠️ Configuration override active';
    envIndicator.appendChild(overrideIndicator);
  }
  
  // Insert after the configuration form if present
  const configForm = document.getElementById('config-form');
  const existingIndicator = document.querySelector('.environment-indicator');
  if (existingIndicator) {
    existingIndicator.remove();
  }
  if (configForm) {
    configForm.appendChild(envIndicator);
  }
}

/**
 * Load saved configuration (for backward compatibility and overrides)
 */
function loadSavedConfiguration() {
  // Load any legacy configuration
  const savedConfig = localStorage.getItem('firebase_config');
  if (savedConfig) {
    try {
      const config = JSON.parse(savedConfig);
      // Convert to new override format if needed
      if (config.djangoUrl || config.firebaseApiKey) {
        saveConfigOverride({
          djangoBackendUrl: config.djangoUrl,
          firebaseApiKey: config.firebaseApiKey,
        });
        // Clean up old format
        localStorage.removeItem('firebase_config');
      }
    } catch (error) {
      console.warn('Failed to load legacy configuration:', error);
    }
  }
}

/**
 * Handle configuration changes
 */
function handleConfigurationChange() {
  const djangoUrl = (document.getElementById('django-url') as HTMLInputElement).value.trim();
  const firebaseApiKey = (document.getElementById('firebase-api-key') as HTMLInputElement).value.trim();
  
  // Get current environment configuration
  const envConfig = getAppConfig();
  
  // Check if user is trying to override environment settings
  const needsOverride = 
    djangoUrl !== envConfig.djangoBackendUrl || 
    firebaseApiKey !== envConfig.firebaseConfig.apiKey;
  
  if (needsOverride) {
    // Save as override configuration
    const override = {
      djangoBackendUrl: djangoUrl || undefined,
      firebaseApiKey: firebaseApiKey || undefined,
    };
    
    saveConfigOverride(override);
    
    // Update auth manager with new settings
    if (djangoUrl) {
      authManager.updateBaseUrl(djangoUrl);
    }
    
    if (firebaseApiKey) {
      authManager.setFirebaseConfig(firebaseApiKey);
    }
    
    // Update UI to show override
    updateConfigurationUI();
    
    showNotification('Configuration override saved. Using custom settings.', false);
  } else {
    // Remove override if values match environment
    clearConfigOverride();
    updateConfigurationUI();
    showNotification('Using environment configuration.', false);
  }
}

/**
 * Show register form
 */
function showRegisterForm() {
  document.getElementById('login-form')?.classList.add('hidden');
  document.getElementById('register-form')?.classList.remove('hidden');
}

/**
 * Show login form
 */
function showLoginForm() {
  document.getElementById('register-form')?.classList.add('hidden');
  document.getElementById('login-form')?.classList.remove('hidden');
}

/**
 * Handle login button click
 */
async function handleLogin() {
  const email = (document.getElementById('auth-email') as HTMLInputElement).value;
  const password = (document.getElementById('auth-password') as HTMLInputElement).value;
  
  if (!email || !password) {
    showAuthStatus('Please enter email and password', true);
    return;
  }
  
  if (!validateConfiguration()) {
    return;
  }
  
  showAuthStatus('Logging in...', false);
  
  const credentials: AuthCredentials = { email, password };
  const result = await authManager.login(credentials);
  
  if (result.success) {
    showAuthenticatedState(result.firebaseUser!, result.user);
    await loadConnections();
    showAuthStatus('Login successful!', false);
    
    // Hide auth status after success
    setTimeout(() => {
      const statusElement = document.getElementById('auth-status');
      if (statusElement) {
        statusElement.style.display = 'none';
      }
    }, 2000);
  } else {
    showAuthStatus(result.message, true);
  }
}

/**
 * Handle registration button click
 */
async function handleRegister() {
  const firstName = (document.getElementById('register-first-name') as HTMLInputElement).value;
  const lastName = (document.getElementById('register-last-name') as HTMLInputElement).value;
  const email = (document.getElementById('register-email') as HTMLInputElement).value;
  const password = (document.getElementById('register-password') as HTMLInputElement).value;
  const confirmPassword = (document.getElementById('register-confirm-password') as HTMLInputElement).value;
  
  if (!firstName || !lastName || !email || !password || !confirmPassword) {
    showAuthStatus('Please fill in all fields', true);
    return;
  }
  
  if (password !== confirmPassword) {
    showAuthStatus('Passwords do not match', true);
    return;
  }
  
  if (password.length < 6) {
    showAuthStatus('Password must be at least 6 characters long', true);
    return;
  }
  
  if (!validateConfiguration()) {
    return;
  }
  
  showAuthStatus('Creating account...', false);
  
  const credentials: RegisterCredentials = {
    email,
    password,
    firstName,
    lastName,
  };
  
  const result = await authManager.register(credentials);
  
  if (result.success) {
    showAuthStatus(result.message, false);
    
    // Clear form fields
    (document.getElementById('register-first-name') as HTMLInputElement).value = '';
    (document.getElementById('register-last-name') as HTMLInputElement).value = '';
    (document.getElementById('register-email') as HTMLInputElement).value = '';
    (document.getElementById('register-password') as HTMLInputElement).value = '';
    (document.getElementById('register-confirm-password') as HTMLInputElement).value = '';
    
    // Switch back to login form
    setTimeout(() => {
      showLoginForm();
    }, 3000);
  } else {
    showAuthStatus(result.message, true);
  }
}

/**
 * Handle refresh user info button click
 */
async function handleRefreshUserInfo() {
  if (!authManager || !authManager.isAuthenticated()) {
    showAuthStatus('Not authenticated', true);
    return;
  }
  
  showAuthStatus('Refreshing user info...', false);
  
  try {
    const user = await authManager.getUserInfo();
    if (user) {
      updateUserInfoDisplay(user);
      showAuthStatus('User info refreshed successfully', false);
    } else {
      showAuthStatus('Failed to refresh user info', true);
    }
  } catch (error) {
    showAuthStatus(`Error refreshing user info: ${error.message}`, true);
  }
}

/**
 * Validate configuration before authentication
 */
function validateConfiguration(): boolean {
  const validation = validateConfig();
  if (!validation.isValid) {
    showAuthStatus(`Configuration error: ${validation.errors.join(', ')}`, true);
  }
  return validation.isValid;
}

/**
 * Handle logout button click
 */
async function handleLogout() {
  await authManager.logout();
  showUnauthenticatedState();
  showAuthStatus('Logged out successfully', false);
}

/**
 * Show authenticated state in UI
 */
function showAuthenticatedState(firebaseUser: FirebaseUser, djangoUser?: UserInfo) {
  // Hide login/register forms, show user info
  const loginForm = document.getElementById('login-form');
  const registerForm = document.getElementById('register-form');
  const userInfo = document.getElementById('user-info');
  const olapSection = document.getElementById('olap-section');
  const operationsSection = document.getElementById('operations-section');
  
  if (loginForm) loginForm.classList.add('hidden');
  if (registerForm) registerForm.classList.add('hidden');
  if (userInfo) userInfo.classList.remove('hidden');
  if (olapSection) olapSection.classList.remove('hidden');
  if (operationsSection) operationsSection.classList.remove('hidden');
  
  // Update user info display
  updateUserInfoDisplay(djangoUser, firebaseUser);
}

/**
 * Update user info display
 */
function updateUserInfoDisplay(djangoUser?: UserInfo, firebaseUser?: FirebaseUser) {
  if (!firebaseUser) {
    firebaseUser = authManager.getFirebaseUser();
  }
  
  const usernameDisplay = document.getElementById('username-display');
  const userEmail = document.getElementById('user-email');
  const emailVerified = document.getElementById('email-verified');
  const creditBalance = document.getElementById('credit-balance');
  
  if (usernameDisplay && firebaseUser) {
    usernameDisplay.textContent = firebaseUser.displayName || firebaseUser.email.split('@')[0];
  }
  
  if (userEmail && firebaseUser) {
    userEmail.textContent = firebaseUser.email;
  }
  
  if (emailVerified && firebaseUser) {
    emailVerified.textContent = firebaseUser.emailVerified ? 'Yes' : 'No';
    emailVerified.className = firebaseUser.emailVerified ? 'email-verified' : 'email-not-verified';
  }
  
  if (creditBalance && djangoUser) {
    creditBalance.textContent = djangoUser.credit_balance.toString();
  }
}

/**
 * Show unauthenticated state in UI
 */
function showUnauthenticatedState() {
  // Show login form, hide authenticated sections
  const loginForm = document.getElementById('login-form');
  const registerForm = document.getElementById('register-form');
  const userInfo = document.getElementById('user-info');
  const olapSection = document.getElementById('olap-section');
  const operationsSection = document.getElementById('operations-section');
  
  if (loginForm) loginForm.classList.remove('hidden');
  if (registerForm) registerForm.classList.add('hidden');
  if (userInfo) userInfo.classList.add('hidden');
  if (olapSection) olapSection.classList.add('hidden');
  if (operationsSection) operationsSection.classList.add('hidden');
  
  // Clear form fields
  const authEmail = document.getElementById('auth-email') as HTMLInputElement;
  const authPassword = document.getElementById('auth-password') as HTMLInputElement;
  
  if (authEmail) authEmail.value = '';
  if (authPassword) authPassword.value = '';
}

/**
 * Show authentication status message
 */
function showAuthStatus(message: string, isError: boolean) {
  const statusElement = document.getElementById('auth-status');
  const textElement = statusElement?.querySelector('.ms-MessageBar-text');
  
  if (statusElement && textElement) {
    textElement.textContent = message;
    statusElement.className = isError ? 
      'ms-MessageBar ms-MessageBar--error' : 
      'ms-MessageBar ms-MessageBar--success';
    statusElement.style.display = 'block';
  }
}

/**
 * Load and display OLAP connections
 */
async function loadConnections() {
  const connections = await authManager.getConnections();
  updateConnectionsList(connections);
  updateActiveConnectionDropdown(connections);
}

/**
 * Update connections list display
 */
function updateConnectionsList(connections: OLAPConnection[]) {
  const connectionsList = document.getElementById('connections-list');
  if (!connectionsList) return;
  
  if (connections.length === 0) {
    connectionsList.innerHTML = '<p class="ms-font-s">No connections saved.</p>';
    return;
  }
  
  const html = connections.map(conn => `
    <div class="connection-item" style="border: 1px solid #ccc; padding: 10px; margin: 5px 0; border-radius: 4px;">
      <div class="ms-font-m"><strong>${conn.name}</strong></div>
      <div class="ms-font-s">Type: ${conn.olap_type.toUpperCase()}</div>
      <div class="ms-font-s">URL: ${conn.server_url}</div>
      <div class="ms-font-s">Status: ${conn.is_active ? 'Active' : 'Inactive'}</div>
      <button onclick="deleteConnection(${conn.id})" class="ms-Button ms-Button--default" style="margin-top: 5px;">
        <span class="ms-Button-label">Delete</span>
      </button>
    </div>
  `).join('');
  
  connectionsList.innerHTML = html;
}

/**
 * Update active connection dropdown
 */
function updateActiveConnectionDropdown(connections: OLAPConnection[]) {
  const dropdown = document.getElementById('active-connection') as HTMLSelectElement;
  if (!dropdown) return;
  
  // Clear existing options except the first one
  dropdown.innerHTML = '<option value="">Select a connection...</option>';
  
  connections.forEach(conn => {
    const option = document.createElement('option');
    option.value = conn.id.toString();
    option.textContent = `${conn.name} (${conn.olap_type.toUpperCase()})`;
    dropdown.appendChild(option);
  });
}

/**
 * Global function to delete connection (called from HTML)
 */
(window as any).deleteConnection = async function(connectionId: number) {
  const result = await authManager.deleteConnection(connectionId);
  if (result.success) {
    await loadConnections();
    showConnectionStatus('Connection deleted successfully', false);
  } else {
    showConnectionStatus(result.error || 'Failed to delete connection', true);
  }
};

/**
 * Save OLAP connection
 */
async function saveConnection() {
  const name = (document.getElementById('connection-name') as HTMLInputElement).value;
  const olapType = (document.getElementById('olap-type') as HTMLSelectElement).value as OLAPConnection['olap_type'];
  const serverUrl = (document.getElementById('server-url') as HTMLInputElement).value;
  const application = (document.getElementById('application') as HTMLInputElement).value;
  const username = (document.getElementById('olap-username') as HTMLInputElement).value;
  const password = (document.getElementById('olap-password') as HTMLInputElement).value;
  const description = ''; // Could add description field later
  
  if (!name || !olapType || !serverUrl || !username || !password) {
    showConnectionStatus('Please fill in all required fields', true);
    return;
  }
  
  showConnectionStatus('Saving connection...', false);
  
  const result = await authManager.createConnection({
    name,
    olap_type: olapType,
    server_url: serverUrl,
    application,
    username,
    password,
    description,
  });
  
  if (result.success) {
    showConnectionStatus('Connection saved successfully!', false);
    await loadConnections();
    
    // Clear form
    (document.getElementById('connection-name') as HTMLInputElement).value = '';
    (document.getElementById('server-url') as HTMLInputElement).value = '';
    (document.getElementById('application') as HTMLInputElement).value = '';
    (document.getElementById('olap-username') as HTMLInputElement).value = '';
    (document.getElementById('olap-password') as HTMLInputElement).value = '';
  } else {
    showConnectionStatus(result.error || 'Failed to save connection', true);
  }
}

/**
 * Show connection status message
 */
function showConnectionStatus(message: string, isError: boolean) {
  const statusElement = document.getElementById('connection-status');
  const textElement = statusElement?.querySelector('.ms-MessageBar-text');
  
  if (statusElement && textElement) {
    textElement.textContent = message;
    statusElement.className = isError ? 
      'ms-MessageBar ms-MessageBar--error' : 
      'ms-MessageBar ms-MessageBar--success';
    statusElement.style.display = 'block';
    
    // Auto-hide after 3 seconds
    setTimeout(() => {
      statusElement.style.display = 'none';
    }, 3000);
  }
}

/**
 * Update chunking strategy based on UI inputs
 */
function updateChunkingStrategy() {
  const maxCells = parseInt((document.getElementById('max-cells-chunk') as HTMLInputElement).value) || 10000;
  const maxSize = parseFloat((document.getElementById('max-payload-size') as HTMLInputElement).value) || 1;
  const chunkByDimension = (document.getElementById('chunk-by-dimension') as HTMLInputElement).checked;
  
  dataProcessor.updateChunkingStrategy({
    maxCellsPerChunk: maxCells,
    maxPayloadSize: maxSize * 1024 * 1024, // Convert MB to bytes
    chunkByDimension: chunkByDimension,
  });
}

/**
 * Handle cost estimation
 */
async function handleEstimateCost() {
  if (!authManager.isAuthenticated()) {
    showNotification('Please login first', true);
    return;
  }
  
  try {
    await Excel.run(async (context) => {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["values"]);
      await context.sync();
      
      const range = selectedRange.values as string[][];
      const estimation = await estimateProcessingCost(range, 0.001); // 0.001 credits per cell
      
      const message = `Estimated: ${estimation.estimatedCells} cells, ${estimation.chunkCount} chunks, ${estimation.estimatedCost.toFixed(4)} credits`;
      showNotification(message, false);
    });
  } catch (error) {
    showNotification(`Estimation failed: ${error.message}`, true);
  }
}

/**
 * Handle refresh data operation
 */
async function handleRefreshData() {
  if (!authManager.isAuthenticated()) {
    showNotification('Please login first', true);
    return;
  }
  
  const connectionId = (document.getElementById('active-connection') as HTMLSelectElement).value;
  if (!connectionId) {
    showNotification('Please select a connection', true);
    return;
  }
  
  if (currentOperation?.isActive) {
    showNotification('Another operation is already in progress', true);
    return;
  }
  
  try {
    await Excel.run(async (context) => {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["values", "address"]);
      await context.sync();
      
      const range = selectedRange.values as string[][];
      const rangeAddress = selectedRange.address;
      
      // Step 1: Preprocess data
      showNotification('Preprocessing data...', false);
      const preprocessed = await dataProcessor.preprocessData(range);
      
      // Step 2: Initialize operation tracking
      currentOperation = {
        isActive: true,
        operationId: `op_${Date.now()}`,
        totalChunks: preprocessed.chunks.length,
        completedChunks: 0,
        currentChunk: 0,
        estimatedCells: preprocessed.metadata.estimatedTotalCells,
        estimatedCost: preprocessed.metadata.estimatedTotalCells * 0.001,
        startTime: new Date(),
        canCancel: true,
      };
      
      showProgressSection(true);
      updateProgress();
      
      // Step 3: Process chunks
      const responses: ChunkResponse[] = [];
      
      for (let i = 0; i < preprocessed.chunks.length; i++) {
        if (!currentOperation?.isActive) {
          // Operation was cancelled
          break;
        }
        
        currentOperation.currentChunk = i + 1;
        updateProgress();
        
        const chunk = preprocessed.chunks[i];
        
        try {
          const response = await processChunk(chunk, parseInt(connectionId));
          responses.push({
            chunkIndex: i,
            chunkId: chunk.chunkId,
            status: 'success',
            data: response,
            actualCells: response.rows?.reduce((sum, row) => sum + row.data.length, 0) || 0,
          });
          
          currentOperation.completedChunks++;
          
        } catch (error) {
          responses.push({
            chunkIndex: i,
            chunkId: chunk.chunkId,
            status: 'error',
            error: error.message,
          });
        }
      }
      
      if (currentOperation?.isActive) {
        // Step 4: Parse and map responses
        showNotification('Mapping data to Excel...', false);
        
        const result = await responseParser.parseAndMapResponse(
          responses,
          preprocessed.structure,
          selectedRange
        );
        
        if (result.success) {
          showNotification(`Data refreshed: ${result.cellsUpdated} cells updated in ${result.rangesUpdated} ranges`, false);
        } else {
          showNotification(`Refresh completed with errors: ${result.errors.join(', ')}`, true);
        }
      }
      
      // Reset operation
      currentOperation = null;
      showProgressSection(false);
    });
    
  } catch (error) {
    showNotification(`Refresh failed: ${error.message}`, true);
    currentOperation = null;
    showProgressSection(false);
  }
}

/**
 * Process a single chunk
 */
async function processChunk(chunk: any, connectionId: number): Promise<any> {
  const response = await authManager.makeAuthenticatedRequest('/api/olap/export-data/', {
    method: 'POST',
    body: JSON.stringify({
      connection_id: connectionId,
      chunk_data: chunk.gridDefinition,
      chunk_metadata: chunk.metadata,
    }),
  });
  
  if (!response.success) {
    throw new Error(response.error || 'Chunk processing failed');
  }
  
  return response.data;
}

/**
 * Handle operation cancellation
 */
function handleCancelOperation() {
  if (currentOperation?.isActive) {
    currentOperation.isActive = false;
    currentOperation.canCancel = false;
    showNotification('Operation cancelled', false);
    showProgressSection(false);
    currentOperation = null;
  }
}

/**
 * Show/hide progress section
 */
function showProgressSection(show: boolean) {
  const progressSection = document.getElementById('progress-section');
  if (progressSection) {
    if (show) {
      progressSection.classList.remove('hidden');
    } else {
      progressSection.classList.add('hidden');
    }
  }
}

/**
 * Update progress display
 */
function updateProgress() {
  if (!currentOperation) return;
  
  const progressBar = document.getElementById('progress-bar');
  const progressText = document.getElementById('progress-text');
  const cancelButton = document.getElementById('cancel-operation') as HTMLButtonElement;
  
  const percentage = (currentOperation.completedChunks / currentOperation.totalChunks) * 100;
  
  if (progressBar) {
    // Remove all progress classes
    progressBar.className = progressBar.className.replace(/progress-\d+/g, '');
    // Add new progress class
    const progressClass = `progress-${Math.round(percentage / 10) * 10}`;
    progressBar.classList.add(progressClass);
  }
  
  if (progressText) {
    progressText.textContent = `${currentOperation.completedChunks} of ${currentOperation.totalChunks} chunks completed`;
  }
  
  if (cancelButton) {
    cancelButton.disabled = !currentOperation.canCancel;
  }
}

function loadSavedSettings() {
  const savedSettings = localStorage.getItem('epmSettings');
  if (savedSettings) {
    const settings: EPMSettings = JSON.parse(savedSettings);
    (document.getElementById('server-url') as HTMLInputElement).value = settings.serverUrl || '';
    (document.getElementById('application') as HTMLInputElement).value = settings.application || '';
    (document.getElementById('olap-username') as HTMLInputElement).value = settings.username || '';
    (document.getElementById('olap-password') as HTMLInputElement).value = settings.password || '';
  }
}

/**
 * Show notification message
 */
function showNotification(message: string, isError: boolean = false) {
  // Create or update notification element
  let notificationElement = document.getElementById('notification-area');
  if (!notificationElement) {
    notificationElement = document.createElement('div');
    notificationElement.id = 'notification-area';
    const appBody = document.getElementById('app-body');
    if (appBody) {
      appBody.appendChild(notificationElement);
    }
  }
  
  const notification = document.createElement('div');
  notification.className = isError ? 
    'ms-MessageBar ms-MessageBar--error' : 
    'ms-MessageBar ms-MessageBar--success';
  notification.innerHTML = `<span class="ms-MessageBar-text">${message}</span>`;
  notification.style.marginBottom = '10px';
  
  notificationElement.appendChild(notification);
  
  // Auto-remove after 5 seconds
  setTimeout(() => {
    if (notification.parentNode) {
      notification.parentNode.removeChild(notification);
    }
  }, 5000);
}

function saveSettings() {
  const settings: EPMSettings = {
    serverUrl: (document.getElementById('server-url') as HTMLInputElement).value,
    application: (document.getElementById('application') as HTMLInputElement).value,
    username: (document.getElementById('olap-username') as HTMLInputElement).value,
    password: (document.getElementById('olap-password') as HTMLInputElement).value
  };

  // Save to localStorage
  localStorage.setItem('epmSettings', JSON.stringify(settings));

  showNotification('Settings saved successfully!', false);
}

// Export settings getter for use in functions.ts (legacy support)
export function getEPMSettings(): EPMSettings {
  const savedSettings = localStorage.getItem('epmSettings');
  if (!savedSettings) {
    throw new Error('EPM settings not configured. Please configure settings in the taskpane.');
  }
  return JSON.parse(savedSettings);
}

// Export auth manager for functions.ts
export function getAuthManagerInstance(): AuthManager {
  return authManager;
}

// Export data processor for functions.ts
export function getDataProcessorInstance(): DataPreprocessor {
  return dataProcessor;
}

// Export response parser for functions.ts
export function getResponseParserInstance(): ResponseParser {
  return responseParser;
}

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "yellow";
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

// Legacy function for backward compatibility
export async function refreshAdhocData(cubeName?: string) {
  // Redirect to new enhanced refresh function
  if (authManager?.isAuthenticated()) {
    await handleRefreshData();
      } else {
    showNotification('Please login to Django backend first', true);
  }
}
