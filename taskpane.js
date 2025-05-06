// Flingoos Excel Logger
// Captures Excel events and sends them to a local listener server

(function() {
    'use strict';

    // Configuration
    const LISTENER_URL = 'https://localhost:5555/log';
    const EVENT_BATCH_INTERVAL = 1000; // milliseconds
    // const MAX_RETRIES = 10; // Removed - no longer needed with infinite retry
    const CONTENT_THROTTLE_TIME = 200; // milliseconds throttle for content capture
    const MAX_LOG_ENTRIES = 5; // Maximum number of log entries to display
    
    // Retry strategy configuration
    const RETRY_TIER_1_INTERVAL = 1000; // 1 second
    const RETRY_TIER_1_DURATION = 15 * 60 * 1000; // 15 minutes
    const RETRY_TIER_2_INTERVAL = 3000; // 3 seconds
    const RETRY_TIER_2_DURATION = 15 * 60 * 1000; // 15 minutes
    const RETRY_TIER_3_INTERVAL = 60000; // 1 minute

    // State
    let isLogging = false;
    let eventQueue = [];
    let batchInterval = null;
    let connectionStatus = 'disconnected';
    let retryCount = 0;
    let retryStartTime = 0;
    let lastContentCapture = 0;

    // Initialize when Office is ready
    Office.onReady(function(info) {
        if (info.host === Office.HostType.Excel) {
            // We no longer need these button handlers since we auto-start
            // document.getElementById('start-logging').onclick = startLogging;
            // document.getElementById('stop-logging').onclick = stopLogging;
            
            // Check if server is available
            checkServerConnection();
            
            // Log initialization
            appendToLog('Add-in initialized');
        }
    });

    // Check if the listener server is available
    function checkServerConnection() {
        console.log('Checking connection to HTTPS server...');
        fetch(LISTENER_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                event_type: 'connection_check',
                timestamp: Date.now(),
                data: { status: 'checking' }
            })
        })
        .then(response => {
            console.log('Server response:', response.status);
            if (response.ok) {
                updateConnectionStatus('connected');
                retryCount = 0;
                // Auto-start logging when connection is established
                startLogging();
            } else {
                console.log('Server returned error status:', response.status);
                updateConnectionStatus('disconnected');
                retryWithBackoff();
            }
        })
        .catch(error => {
            console.log('Connection error details:', error.message, error.name);
            updateConnectionStatus('disconnected');
            retryWithBackoff();
        });
    }

    // Retry connection with tiered strategy
    function retryWithBackoff() {
        const now = Date.now();
        
        // Initialize retry start time if this is the first retry
        if (retryCount === 0) {
            retryStartTime = now;
        }
        
        retryCount++;
        const elapsedTime = now - retryStartTime;
        let nextRetryInterval;
        let retryTier;
        
        // Tier 1: Every second for first 15 minutes
        if (elapsedTime < RETRY_TIER_1_DURATION) {
            nextRetryInterval = RETRY_TIER_1_INTERVAL;
            retryTier = 1;
        }
        // Tier 2: Every 3 seconds for next 15 minutes
        else if (elapsedTime < RETRY_TIER_1_DURATION + RETRY_TIER_2_DURATION) {
            nextRetryInterval = RETRY_TIER_2_INTERVAL;
            retryTier = 2;
        }
        // Tier 3: Every minute indefinitely
        else {
            nextRetryInterval = RETRY_TIER_3_INTERVAL;
            retryTier = 3;
        }
        
        appendToLog(`Connection attempt #${retryCount} failed. Tier ${retryTier}: Retrying in ${nextRetryInterval/1000} seconds...`);
        setTimeout(checkServerConnection, nextRetryInterval);
    }

    // Update the connection status UI
    function updateConnectionStatus(status) {
        connectionStatus = status;
        const statusElement = document.getElementById('status');
        statusElement.className = `status ${status}`;
        
        if (status === 'connected') {
            statusElement.innerHTML = `<div class="status-dot"></div><div>Status: Connected</div>`;
            retryCount = 0; // Reset retry count when connected
            retryStartTime = 0;
            appendToLog(`Connection status: ${status}`);
        } else {
            // Show retry information in the status display
            if (retryCount > 0) {
                const elapsedTime = Math.floor((Date.now() - retryStartTime) / 1000);
                let retryTier;
                
                if (elapsedTime < RETRY_TIER_1_DURATION / 1000) {
                    retryTier = 1;
                } else if (elapsedTime < (RETRY_TIER_1_DURATION + RETRY_TIER_2_DURATION) / 1000) {
                    retryTier = 2;
                } else {
                    retryTier = 3;
                }
                
                statusElement.innerHTML = `<div class="status-dot"></div><div>Status: Disconnected (Retry #${retryCount}, Tier ${retryTier})</div>`;
            } else {
                statusElement.innerHTML = `<div class="status-dot"></div><div>Status: Disconnected</div>`;
            }
            appendToLog(`Connection status: ${status}`);
        }
    }
    
    // Start logging Excel events
    function startLogging() {
        if (connectionStatus !== 'connected') {
            appendToLog('Cannot start logging: not connected to server');
            return;
        }
        
        if (isLogging) {
            // Already logging
            return;
        }
        
        isLogging = true;
        
        // Update UI
        // document.getElementById('start-logging').disabled = true;
        // document.getElementById('stop-logging').disabled = false;
        appendToLog('Started logging Excel events');
        
        // Register event handlers
        registerExcelEventHandlers();
        
        // Start batch sending interval
        batchInterval = setInterval(sendEventBatch, EVENT_BATCH_INTERVAL);
        
        // Log initial workbook state
        logWorkbookInfo();
    }
    
    // Stop logging Excel events
    function stopLogging() {
        if (!isLogging) {
            return;
        }
        
        isLogging = false;
        
        // Update UI
        // document.getElementById('start-logging').disabled = (connectionStatus !== 'connected');
        // document.getElementById('stop-logging').disabled = true;
        appendToLog('Stopped logging Excel events');
        
        // Unregister event handlers
        unregisterExcelEventHandlers();
        
        // Stop batch sending
        if (batchInterval) {
            clearInterval(batchInterval);
            batchInterval = null;
        }
        
        // Send any remaining events
        if (eventQueue.length > 0) {
            sendEventBatch();
        }
    }
    
    // Register all Excel event handlers
    function registerExcelEventHandlers() {
        Excel.run(async (context) => {
            // Get the current worksheet
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Selection change events
            sheet.onSelectionChanged.add(handleSelectionChange);
            
            // Worksheet events
            sheet.onActivated.add(handleSheetActivated);
            sheet.onDeactivated.add(handleSheetDeactivated);
            
            // Workbook events - wrapped in try/catch since onSaved might not be available
            try {
                if (context.workbook.onSaved) {
                    context.workbook.onSaved.add(handleWorkbookSaved);
                } else {
                    console.log("workbook.onSaved event not available in this Excel version");
                }
            } catch (e) {
                console.log("Unable to register workbook.onSaved handler:", e.message);
            }
            
            await context.sync();

            // Capture initial selection content and log debug info
            console.log("About to capture initial cell content");
            captureActiveSelectionContent(context);
        }).catch(handleError);
        
        // We need to use event listeners for data changed events
        // since they can't be registered through the Excel.run context
        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            handleDocumentSelectionChanged
        );
        
        Office.context.document.addHandlerAsync(
            "documentSelectionChanged", 
            handleDocumentSelectionChanged
        );
        
        // Since there's no direct event for cell edits in Office.js,
        // we'll track selection changes and check for changes
        setInterval(checkForCellChanges, 1000);
    }
    
    // Track the last selection to detect edits
    let lastSelection = null;
    let lastCellValues = {};
    
    // Check for cell changes by comparing values
    function checkForCellChanges() {
        if (!isLogging || !lastSelection) return;
        
        Excel.run(async (context) => {
            // Get the current selection
            const range = context.workbook.getSelectedRange();
            range.load('address,values,formulas,numberFormat');
            
            await context.sync();
            
            // If we have lastCellValues, compare to detect changes
            if (Object.keys(lastCellValues).length > 0) {
                const currentAddress = range.address;
                const currentValues = range.values;
                const currentFormulas = range.formulas;
                
                // Check if we have values for this range
                if (lastCellValues[currentAddress]) {
                    const previousValues = lastCellValues[currentAddress].values;
                    const previousFormulas = lastCellValues[currentAddress].formulas;
                    
                    // Compare values to detect changes
                    let hasChanges = false;
                    let changedCells = [];
                    
                    // Simple matrix comparison
                    if (previousValues && currentValues && 
                        previousValues.length === currentValues.length) {
                        
                        for (let i = 0; i < currentValues.length; i++) {
                            if (previousValues[i].length === currentValues[i].length) {
                                for (let j = 0; j < currentValues[i].length; j++) {
                                    if (previousValues[i][j] !== currentValues[i][j] ||
                                        (previousFormulas && currentFormulas && 
                                         previousFormulas[i][j] !== currentFormulas[i][j])) {
                                        
                                        hasChanges = true;
                                        changedCells.push({
                                            row: i,
                                            column: j,
                                            previousValue: previousValues[i][j],
                                            currentValue: currentValues[i][j],
                                            previousFormula: previousFormulas ? previousFormulas[i][j] : null,
                                            currentFormula: currentFormulas ? currentFormulas[i][j] : null
                                        });
                                    }
                                }
                            }
                        }
                    }
                    
                    // If changes detected, log them
                    if (hasChanges) {
                        queueEvent('cell_edit', {
                            address: currentAddress,
                            changes: changedCells,
                            is_formula: currentFormulas.some(row => 
                                row.some(cell => typeof cell === 'string' && cell.startsWith('='))
                            ),
                            current_values: currentValues,
                            current_formulas: currentFormulas
                        });

                        // Also capture full content after edit
                        captureContentIfNeeded(range);
                    }
                }
                
                // Update stored values for this range
                lastCellValues[currentAddress] = {
                    values: currentValues,
                    formulas: currentFormulas
                };
            } else {
                // Initialize lastCellValues if empty
                lastCellValues[range.address] = {
                    values: range.values,
                    formulas: range.formulas
                };
            }
            
        }).catch(handleError);
    }
    
    // Unregister Excel event handlers
    function unregisterExcelEventHandlers() {
        Excel.run(async (context) => {
            // Get the active worksheet
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // Remove event handlers
            sheet.onSelectionChanged.remove();
            sheet.onActivated.remove();
            sheet.onDeactivated.remove();
            context.workbook.onSaved.remove();
            
            await context.sync();
        }).catch(handleError);
        
        // Remove document event handlers
        Office.context.document.removeHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            { handler: handleDocumentSelectionChanged }
        );
        
        Office.context.document.removeHandlerAsync(
            "documentSelectionChanged", 
            { handler: handleDocumentSelectionChanged }
        );
    }
    
    // Event Handlers
    function handleSelectionChange(event) {
        Excel.run(async (context) => {
            // Get information about the selected range
            const range = context.workbook.getSelectedRange();
            range.load('address,columnCount,rowCount,values,formulas,numberFormat');
            
            await context.sync();
            
            // Save last selection for cell edit detection
            lastSelection = {
                address: range.address,
                values: range.values,
                formulas: range.formulas
            };
            
            // Store values for this range to detect edits later
            lastCellValues[range.address] = {
                values: range.values,
                formulas: range.formulas
            };
            
            queueEvent('selection_changed', {
                address: range.address,
                columns: range.columnCount,
                rows: range.rowCount,
                cellCount: range.columnCount * range.rowCount
            });

            // Capture cell content with throttling
            captureContentIfNeeded(range);
        }).catch(handleError);
    }
    
    function handleDocumentSelectionChanged(eventArgs) {
        if (!isLogging) return;
        
        Excel.run(async (context) => {
            // Get detailed selection information
            const range = context.workbook.getSelectedRange();
            range.load('address,columnCount,rowCount');
            
            await context.sync();
            
            queueEvent('document_selection_changed', {
                address: range.address,
                columns: range.columnCount,
                rows: range.rowCount
            });
        }).catch(handleError);
    }
    
    function handleSheetActivated(event) {
        Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(event.worksheetId);
            sheet.load('name, id, position');
            
            await context.sync();
            
            queueEvent('sheet_activated', {
                name: sheet.name,
                id: sheet.id,
                position: sheet.position
            });

            // Capture active selection content when sheet is activated
            captureActiveSelectionContent(context);
        }).catch(handleError);
    }
    
    function handleSheetDeactivated(event) {
        queueEvent('sheet_deactivated', {
            worksheetId: event.worksheetId
        });
    }
    
    function handleWorkbookSaved(event) {
        queueEvent('workbook_saved', {});
    }
    
    // Log initial workbook information
    function logWorkbookInfo() {
        Excel.run(async (context) => {
            // Get workbook details
            const workbook = context.workbook;
            workbook.load('name');
            
            // Get all worksheets
            const worksheets = workbook.worksheets;
            worksheets.load('items/name,items/position,items/visibility');
            
            // Get active sheet
            const activeSheet = workbook.worksheets.getActiveWorksheet();
            activeSheet.load('name, id, position');
            
            // Get selected range
            const selection = workbook.getSelectedRange();
            selection.load('address,columnCount,rowCount');
            
            await context.sync();
            
            // Log workbook info
            queueEvent('workbook_info', {
                name: workbook.name,
                worksheets: worksheets.items.map(sheet => ({
                    name: sheet.name,
                    position: sheet.position,
                    visibility: sheet.visibility
                })),
                activeSheet: {
                    name: activeSheet.name,
                    id: activeSheet.id,
                    position: activeSheet.position
                },
                selection: {
                    address: selection.address,
                    columns: selection.columnCount,
                    rows: selection.rowCount
                }
            });
        }).catch(handleError);
    }
    
    // Queue an event to be sent in the next batch
    function queueEvent(eventType, data) {
        if (!isLogging) return;
        
        const event = {
            event_type: eventType,
            timestamp: Date.now(),
            data: data
        };
        
        eventQueue.push(event);
        appendToLog(`Event queued: ${eventType}`);
    }
    
    // Send a batch of events to the listener server
    function sendEventBatch() {
        if (eventQueue.length === 0) return;
        
        // Create a copy of the queue and clear it
        const batch = [...eventQueue];
        eventQueue = [];
        
        fetch(LISTENER_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                event_type: 'batch',
                timestamp: Date.now(),
                events: batch
            })
        })
        .then(response => {
            if (!response.ok) {
                // If failed, add events back to the queue
                eventQueue = [...batch, ...eventQueue];
                appendToLog(`Failed to send batch: ${response.status}`);
                updateConnectionStatus('disconnected');
                checkServerConnection();
            } else {
                appendToLog(`Sent batch of ${batch.length} events`);
            }
        })
        .catch(error => {
            // If error, add events back to the queue
            eventQueue = [...batch, ...eventQueue];
            appendToLog(`Error sending events: ${error.message}`);
            updateConnectionStatus('disconnected');
            checkServerConnection();
        });
    }
    
    // Handle Excel.run errors
    function handleError(error) {
        appendToLog(`Error: ${error.message}`);
        console.error(error);
    }
    
    // Append a message to the log display
    function appendToLog(message) {
        console.log(message); // Still log to console
        
        // Only show important messages in the UI log
        if (message.includes("error") || 
            message.includes("failed") || 
            message.includes("Connected") || 
            message.includes("Disconnected") ||
            message.includes("Started logging") ||
            message.includes("Stopped logging")) {
            
            const logElement = document.getElementById('log');
            const now = new Date();
            const timestamp = now.toLocaleTimeString();
            const logEntry = document.createElement('div');
            logEntry.textContent = `[${timestamp}] ${message}`;
            logElement.appendChild(logEntry);
            
            // Limit the number of visible log entries
            while (logElement.children.length > MAX_LOG_ENTRIES) {
                logElement.removeChild(logElement.firstChild);
            }
            
            // Auto-scroll to bottom
            logElement.scrollTop = logElement.scrollHeight;
        }
    }

    // Capture cell content with throttling
    function captureContentIfNeeded(range) {
        const now = Date.now();
        if (now - lastContentCapture > CONTENT_THROTTLE_TIME) {
            lastContentCapture = now;
            
            // If range is already loaded, capture content directly
            if (range && range.values) {
                captureCellContent(range);
            } else {
                // Otherwise get active selection
                captureActiveSelectionContent();
            }
        }
    }
    
    // Capture content of active selection
    function captureActiveSelectionContent(context) {
        Excel.run(async (ctx) => {
            // Use the provided context or create a new one
            const runContext = context || ctx;
            
            // Get the current selection
            const range = runContext.workbook.getSelectedRange();
            range.load(['address', 'values', 'formulas', 'numberFormat', 'rowCount', 'columnCount']);
            
            await runContext.sync();
            
            captureCellContent(range);
        }).catch(handleError);
    }
    
    // Process and queue cell content event
    function captureCellContent(range) {
        if (!isLogging) return;
        
        // Debug logging
        console.log("captureCellContent called for range:", range.address);
        appendToLog(`Capturing content for range: ${range.address}`);
        
        // Skip if range is too large (more than 100 cells)
        if (range.rowCount * range.columnCount > 100) {
            // For large ranges, just capture dimensions and a sample
            const sampleValues = range.values.slice(0, 3).map(row => row.slice(0, 3));
            const sampleFormulas = range.formulas.slice(0, 3).map(row => row.slice(0, 3));
            
            appendToLog(`Queueing large range content: ${range.rowCount}x${range.columnCount}`);
            queueEvent('cell_content', {
                address: range.address,
                rowCount: range.rowCount,
                columnCount: range.columnCount,
                is_large_range: true,
                sample_values: sampleValues,
                sample_formulas: sampleFormulas
            });
        } else {
            // For smaller ranges, capture everything
            appendToLog(`Queueing content for: ${range.address}, size: ${range.rowCount}x${range.columnCount}`);
            queueEvent('cell_content', {
                address: range.address,
                rowCount: range.rowCount,
                columnCount: range.columnCount,
                values: range.values,
                formulas: range.formulas,
                numberFormat: range.numberFormat
            });
        }
    }
})(); 