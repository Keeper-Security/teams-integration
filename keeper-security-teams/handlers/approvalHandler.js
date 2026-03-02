/**
 * Approval Action Handler
 * 
 * Re-exports from modular approval handlers for backwards compatibility.
 * 
 * Handles Adaptive Card action submissions for:
 * - Approve/Deny record access requests
 * - Approve/Deny folder access requests
 * - Approve/Deny one-time share requests
 * 
 * @see ./approval/ for modular implementation
 */

module.exports = require('./approval');
