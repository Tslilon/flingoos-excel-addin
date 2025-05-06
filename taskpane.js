// Flingoos Excel Logger
// Captures Excel events and sends them to a local listener server

(function() {
    'use strict';

    // Configuration
    const LISTENER_URL = 'https://localhost:5555/log';
    const EVENT_BATCH_INTERVAL = 1000; // milliseconds
    const MAX_RETRIES = 3;

    // State
    let isLogging = false;
    let eventQueue = [];
    let batchInterval = null;
    let connectionStatus = 'disconnected';
    let retryCount = 0;

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

    // Retry connection with exponential backoff
    function retryWithBackoff() {
        if (retryCount < MAX_RETRIES) {
            const delay = Math.pow(2, retryCount) * 1000;
            retryCount++;
            appendToLog(`Connection failed. Retrying in ${delay/1000} seconds...`);
            setTimeout(checkServerConnection, delay);
        } else {
            appendToLog('Could not connect to listener server. Please check if the server is running.');
        }
    }

    // Update the connection status UI
    function updateConnectionStatus(status) {
        connectionStatus = status;
        const statusElement = document.getElementById('status');
        statusElement.className = `status ${status}`;
        statusElement.innerText = `Status: ${status === 'connected' ? 'Connected' : 'Disconnected'}`;
        
        // We don't need to update buttons anymore since they're hidden
        // document.getElementById('start-logging').disabled = (status !== 'connected' || isLogging);
        // document.getElementById('stop-logging').disabled = !isLogging;
        
        appendToLog(`Connection status: ${status}`);
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
            
            // Workbook events
            context.workbook.onSaved.add(handleWorkbookSaved);
            
            await context.sync();
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
                            )
                        });
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
            range.load('address,columnCount,rowCount,values,formulas');
            
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
        const log = document.getElementById('log');
        const timestamp = new Date().toLocaleTimeString();
        log.innerHTML += `[${timestamp}] ${message}<br>`;
        log.scrollTop = log.scrollHeight;
    }
})(); 