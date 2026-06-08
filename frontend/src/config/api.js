/**
 * API Configuration
 * Points to FastAPI backend
 */

const API_BASE_URL = import.meta.env.VITE_API_URL || "http://localhost:8000";

export const API_ENDPOINTS = {
  STATUS:   `${API_BASE_URL}/api/status`,
  VALIDATE: `${API_BASE_URL}/api/validate`,
  LOAD:     `${API_BASE_URL}/api/load`,
};

export default API_BASE_URL;
