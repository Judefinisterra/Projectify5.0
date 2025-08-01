// Model Update Handler - Handles different processing for new builds vs updates
// Handles encoding changes based on update type

const DEBUG_UPDATE_HANDLER = true;

/**
 * Utility function to strip code block markers from JSON strings
 * @param {string} content - The content that might have code blocks
 * @returns {string} - Cleaned content without code block markers
 */
function stripCodeBlockMarkers(content) {
    if (!content || typeof content !== 'string') {
        return content;
    }
    
    let cleaned = content.trim();
    
    // Remove ```json at the beginning
    if (cleaned.startsWith('```json')) {
        cleaned = cleaned.substring(7).trim();
    }
    
    // Remove ``` at the end
    if (cleaned.endsWith('```')) {
        cleaned = cleaned.substring(0, cleaned.length - 3).trim();
    }
    
    return cleaned;
}

class ModelUpdateHandler {
    constructor() {
        this.isNewBuild = true; // Default to new build
        this.updateData = null;
    }

    /**
     * Determines if this is a new build or an update based on JSON structure
     * @param {Object} jsonData - The JSON data from the AI response
     * @returns {boolean} - true if new build, false if update
     */
    determineUpdateType(jsonData) {
        // Check if JSON has tab_updates structure (indicates update)
        if (jsonData && jsonData.tab_updates) {
            this.isNewBuild = false;
            this.updateData = jsonData.tab_updates;
            if (DEBUG_UPDATE_HANDLER) {
                console.log("[ModelUpdateHandler] Detected UPDATE mode");
                console.log("[ModelUpdateHandler] Update data:", this.updateData);
            }
            return false;
        } else {
            this.isNewBuild = true;
            this.updateData = null;
            if (DEBUG_UPDATE_HANDLER) {
                console.log("[ModelUpdateHandler] Detected NEW BUILD mode");
            }
            return true;
        }
    }

    /**
     * Processes model codes based on update type
     * @param {Array} modelCodes - Array of model codes to process
     * @returns {Array} - Processed model codes
     */
    processModelCodes(modelCodes) {
        if (this.isNewBuild) {
            return this.processNewBuild(modelCodes);
        } else {
            return this.processUpdate(modelCodes);
        }
    }

    /**
     * Processes codes for new build - no modifications needed
     * @param {Array} modelCodes - Array of model codes
     * @returns {Array} - Unmodified model codes
     */
    processNewBuild(modelCodes) {
        if (DEBUG_UPDATE_HANDLER) {
            console.log("[ModelUpdateHandler] Processing new build - no code modifications");
        }
        return modelCodes;
    }

    /**
     * Processes codes for updates - modifies encoding based on update type
     * @param {Array} modelCodes - Array of model codes
     * @returns {Array} - Modified model codes
     */
    processUpdate(modelCodes) {
        if (DEBUG_UPDATE_HANDLER) {
            console.log("[ModelUpdateHandler] Processing update - modifying codes based on update types");
        }

        const processedCodes = modelCodes.map(code => {
            return this.transformCodeForUpdate(code);
        });

        return processedCodes;
    }

    /**
     * Transforms individual code based on update type
     * @param {string} code - Individual model code
     * @returns {string} - Transformed code
     */
    transformCodeForUpdate(code) {
        // Extract tab name from TAB codes
        const tabMatch = code.match(/<TAB;\s*label1="([^"]+)"/);
        if (!tabMatch) {
            return code; // Not a TAB code, return unchanged
        }

        const tabName = tabMatch[1];
        const updateInfo = this.updateData[tabName];

        if (!updateInfo) {
            if (DEBUG_UPDATE_HANDLER) {
                console.log(`[ModelUpdateHandler] No update info found for tab: ${tabName}`);
            }
            return code; // No update info, return unchanged
        }

        const updateType = updateInfo.update_type;

        if (DEBUG_UPDATE_HANDLER) {
            console.log(`[ModelUpdateHandler] Transforming code for tab: ${tabName}, type: ${updateType}`);
        }

        // Transform based on update type
        switch (updateType) {
            case 'adding_to_existing_tab':
                // Change <TAB; to <ADDCODES;
                return code.replace(/<TAB;/, '<ADDCODES;');
                
            case 'replacing_existing_tab':
            case 'adding_new_tab':
                // Keep <TAB; unchanged
                return code;
                
            default:
                if (DEBUG_UPDATE_HANDLER) {
                    console.warn(`[ModelUpdateHandler] Unknown update type: ${updateType}`);
                }
                return code;
        }
    }

    /**
     * Determines which sheets to move to workbook
     * @returns {Array} - Array of sheet names to move
     */
    getSheetsToMove() {
        if (this.isNewBuild) {
            // New builds: Move financials, calcs, codes, misc, actuals
            return ['Financials', 'Calcs', 'Codes', 'Misc', 'Actuals'];
        } else {
            // Updates: Only move codes and calcs
            return ['Codes', 'Calcs'];
        }
    }

    /**
     * Determines whether to delete calcs based on update types
     * @returns {boolean} - true if calcs should be deleted
     */
    shouldDeleteCalcs() {
        if (this.isNewBuild) {
            return false; // Never delete calcs for new builds
        }

        // For updates, check if any tabs are adding_to_existing_tab or replacing_existing_tab
        const updateTypes = Object.values(this.updateData).map(update => update.update_type);
        const shouldDelete = updateTypes.some(type => 
            type === 'adding_to_existing_tab' || 
            type === 'replacing_existing_tab'
        );

        if (DEBUG_UPDATE_HANDLER) {
            console.log(`[ModelUpdateHandler] Should delete calcs: ${shouldDelete}`);
            console.log(`[ModelUpdateHandler] Update types found:`, updateTypes);
        }

        return shouldDelete;
    }

    /**
     * Gets summary of update actions
     * @returns {Object} - Summary of what will be processed
     */
    getUpdateSummary() {
        return {
            isNewBuild: this.isNewBuild,
            sheetsToMove: this.getSheetsToMove(),
            shouldDeleteCalcs: this.shouldDeleteCalcs(),
            updateData: this.updateData,
            tabCount: this.updateData ? Object.keys(this.updateData).length : 0
        };
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { ModelUpdateHandler, stripCodeBlockMarkers };
} else if (typeof window !== 'undefined') {
    window.ModelUpdateHandler = ModelUpdateHandler;
    window.stripCodeBlockMarkers = stripCodeBlockMarkers;
}