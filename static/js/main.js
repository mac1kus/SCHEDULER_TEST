/**
 * Refinery Crude Oil Scheduling System - ENHANCED
 * All JavaScript consolidated in main.js (moved from HTML)
 */

// Global variables
let currentResults = null;

// Configuration objects
const ALERT_TYPES = {
    SUCCESS: 'success',
    WARNING: 'warning',
    DANGER: 'danger',
    INFO: 'info'
};

const TANK_STATUS_COLORS = {
    READY: '#28a745',
    FEEDING: '#28a745',
    SETTLING: '#ffc107',
    LAB_TESTING: '#ffd700',
    FILLING: '#007bff',
    FILLED: '#007bff',
    EMPTY: '#6c757d'
};

const API_ENDPOINTS = {
    SIMULATE: '/api/simulate',
    BUFFER_ANALYSIS: '/api/buffer_analysis',
    CARGO_OPTIMIZATION: '/api/cargo_optimization',
    SAVE_INPUTS: '/api/save_inputs',
    LOAD_INPUTS: '/api/load_inputs',
    EXPORT_DATA: '/api/export_data',
    EXPORT_TANK_STATUS: '/api/export_tank_status'
};

/**
 * Utility Functions
 */
const Utils = {
    formatNumber: (num) => Math.round(num).toLocaleString(),

    showLoading: (show = true) => {
        const loading = document.getElementById('loading');
        if (loading) loading.style.display = show ? 'block' : 'none';
        document.querySelectorAll('button').forEach(btn => btn.disabled = show);
    },

    showResults: () => {
        const results = document.getElementById('results');
        if (results) results.style.display = 'block';
    },

    getTankLevelColor: (volume, deadBottom) => {
        if (volume <= deadBottom) return '#dc3545';
        if (volume < deadBottom * 3) return '#ffc107';
        return '#28a745';
    },

    getStatusColor: (status) => TANK_STATUS_COLORS[status] || '#000',

    createAlert: (type, message) =>
        `<div class="alert alert-${type}">${message}</div>`,

    createMetricCard: (title, value, label, extraContent = '') => `
        <div class="metric-card">
            <h4>${title}</h4>
            <div class="metric-value">${value}</div>
            <div class="metric-label">${label}</div>
            ${extraContent}
        </div>
    `
};

// ===== MOVED FROM HTML =====
// Navigation functions
function scrollToTop() {
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function scrollToCargoReport() {
    const element = document.getElementById('cargoReportContainer');
    if (element) {
        element.scrollIntoView({ behavior: 'smooth' });
    }
}

function scrollToBottom() {
    window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
}

function scrollToSimulation() {
    const element = document.querySelector('.btn-group');
    if (element) {
        element.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
}

// Tank management functions (moved from HTML, using HTML versions)
function updateTankCount() {
    const numTanks = parseInt(document.getElementById('numTanks').value);
    const tankCountDisplay = document.getElementById('tankCountDisplay');
    tankCountDisplay.textContent = `tanks (${numTanks} tanks total)`;
    
    // Update tank grid to show/hide tanks based on count
    const tankGrid = document.getElementById('tankGrid');
    const existingTanks = tankGrid.querySelectorAll('.tank-box').length;
    
    if (numTanks > existingTanks) {
        // Add new tanks
        for (let i = existingTanks + 1; i <= numTanks; i++) {
            addNewTankBox(i);
        }
    } else if (numTanks < existingTanks) {
        // Remove extra tanks
        for (let i = existingTanks; i > numTanks; i--) {
            const tankBox = tankGrid.querySelector(`.tank-box:nth-child(${i})`);
            if (tankBox) {
                tankBox.remove();
            }
        }
    }
}

function addOneTank() {
    const numTanksInput = document.getElementById('numTanks');
    const currentCount = parseInt(numTanksInput.value);
    numTanksInput.value = currentCount + 1;
    updateTankCount();
    autoSaveInputs();
}

function removeOneTank() {
    const numTanksInput = document.getElementById('numTanks');
    const currentCount = parseInt(numTanksInput.value);
    
    // Set a minimum limit to prevent removing all tanks
    const minTanks = 1;

    if (currentCount > minTanks) {
        // Decrease the value of the number input field
        numTanksInput.value = currentCount - 1;

        // Find the last tank's HTML element using its index and remove it
        const tankGrid = document.getElementById('tankGrid');
        const lastTankBox = tankGrid.querySelector(`.tank-box:nth-child(${currentCount})`);
        
        if (lastTankBox) {
            tankGrid.removeChild(lastTankBox);
        }

        // Update the displayed tank count and save the inputs
        updateTankCount();
        autoSaveInputs();
    }
}

function addNewTankBox(tankNumber) {
    const tankGrid = document.getElementById('tankGrid');
    const tankCapacity = document.getElementById('tankCapacity').value;
    
    const tankBox = document.createElement('div');
    tankBox.className = 'tank-box';
    tankBox.innerHTML = `
        <h4>Tank ${tankNumber}</h4>
        <div class="tank-input-row">
            <label>Current Level:</label>
            <input type="number" id="tank${tankNumber}Level" value="0" min="0" max="${tankCapacity}" onchange="autoSaveInputs()">
            <span>bbl</span>
        </div>
        <div class="tank-input-row">
            <label>Dead Bottom:</label>
            <input type="number" id="deadBottom${tankNumber}" value="10000" min="10000" max="10500" onchange="autoSaveInputs()">
            <span>bbl</span>
        </div>
    `;
    
    tankGrid.appendChild(tankBox);
}

/**
 * Get current tank count dynamically
 */
function getCurrentTankCount() {
    const count = parseInt(document.getElementById('numTanks').value);
    // Return the parsed number if it's a valid non-negative integer, otherwise return 0.
    return !isNaN(count) && count >= 0 ? count : 0;
}

/**
 * AUTO-POPULATE TANK LEVELS FROM TANK CAPACITY - Updated for dynamic tanks
 */
function populateTankLevels() {
    const tankCapacity = document.getElementById('tankCapacity').value;
    const numTanks = getCurrentTankCount();

    if (tankCapacity && parseFloat(tankCapacity) > 0) {
        // Populate all active tank levels with tank capacity
        for (let i = 1; i <= numTanks; i++) {
            const tankLevelInput = document.getElementById(`tank${i}Level`);
            if (tankLevelInput) {
                tankLevelInput.value = tankCapacity;
                // Also update the max attribute
                tankLevelInput.setAttribute('max', tankCapacity);
            }
        }
        console.log(`All ${numTanks} tanks populated with ${parseFloat(tankCapacity).toLocaleString()} bbl`);
        // Also validate the inventory range immediately
        validateInventoryRange();
    }
}

/**
 * AUTO-CALCULATE PUMPING DAYS
 */
function autoCalculatePumpingDays() {
    // Get the largest cargo capacity that's enabled
    const vlcc = parseFloat(document.getElementById('vlccCapacity').value) || 0;
    const suezmax = parseFloat(document.getElementById('suezmaxCapacity').value) || 0;
    const aframax = parseFloat(document.getElementById('aframaxCapacity').value) || 0;
    const panamax = parseFloat(document.getElementById('panamaxCapacity').value) || 0;
    const handymax = parseFloat(document.getElementById('handymaxCapacity').value) || 0;

    // Get the largest active cargo
    const largestCargo = Math.max(vlcc, suezmax, aframax, panamax, handymax);
    const pumpingRate = parseFloat(document.getElementById('pumpingRate').value) || 30000;

    if (largestCargo > 0 && pumpingRate > 0) {
        // Calculate pumping hours = cargo size / pumping rate
        const pumpingHours = largestCargo / pumpingRate;
        document.getElementById('pumpingDaysDisplay').value = pumpingHours.toFixed(2);
    } else {
        document.getElementById('pumpingDaysDisplay').value = 0;
    }

    // Also update lead time since pumping hours affects it
    autoCalculateLeadTime();
}

/**
 * AUTO-CALCULATE LEAD TIME
 */
function autoCalculateLeadTime() {
    const preJourney = parseFloat(document.getElementById('preJourneyDays').value) || 0;
    const journey = parseFloat(document.getElementById('journeyDays').value) || 0;
    const preDischarge = parseFloat(document.getElementById('preDischargeDays').value) || 0;
    const settling = parseFloat(document.getElementById('settlingTime').value) || 0;
    const labTesting = parseFloat(document.getElementById('labTestingDays').value) || 0;

    const leadTime = preJourney + journey + preDischarge + settling + labTesting;
    document.getElementById('leadTimeDisplay').value = leadTime.toFixed(1);
}

/**
 * TOGGLE DEPARTURE MODE
 */
function toggleDepartureMode() {
    const mode = document.getElementById('departureMode').value;
    const manualSection = document.getElementById('manualDepartureSection');
    const solverSection = document.getElementById('solverDepartureSection');

    if (mode === 'manual') {
        manualSection.style.display = 'block';
        solverSection.style.display = 'none';
    } else {
        manualSection.style.display = 'none';
        solverSection.style.display = 'block';
    }
}

/**
 * APPLY DEFAULT DEAD BOTTOM - Updated for dynamic tanks
 */
function applyDefaultDeadBottom() {
    const defaultValue = document.getElementById('defaultDeadBottom').value;
    const actualTankCount = document.querySelectorAll('.tank-box').length;
    
    for (let i = 1; i <= actualTankCount; i++) {
        const deadBottomInput = document.getElementById(`deadBottom${i}`);
        if (deadBottomInput) {
            deadBottomInput.value = defaultValue;
        }
    }
    autoSaveInputs();
}

/**
 * COLLECT FORM DATA - Improved for dynamic tanks
 */
function collectFormData() {
    const data = {};

    // Collect all input values more reliably
    document.querySelectorAll('input, select, textarea').forEach(input => {
        if (input.id && input.id !== '') {
            if (input.type === 'checkbox') {
                data[input.id] = input.checked;
            } else if (input.type === 'radio') {
                if (input.checked) {
                    data[input.id] = input.value;
                }
            } else if (input.type === 'number') {
                data[input.id] = parseFloat(input.value) || 0;
            } else {
                data[input.id] = input.value || '';
            }
        }
    });

    return data;
}

/**
 * AUTO-SAVE INPUTS - Improved
 */
async function autoSaveInputs() {
    try {
        const inputs = collectFormData();
        
        // Save to localStorage immediately
        localStorage.setItem('refineryInputs', JSON.stringify(inputs));
        console.log('Inputs saved to localStorage');
        
        // Try to save to backend (don't block if it fails)
        try {
            const response = await fetch(API_ENDPOINTS.SAVE_INPUTS, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(inputs)
            });
            
            if (response.ok) {
                console.log('Inputs saved to server');
                // Optional: Show a brief success indicator
                showSaveStatus('saved');
            } else {
                console.log('Server save failed, but localStorage saved');
            }
        } catch (serverError) {
            console.log('Server unavailable, but localStorage saved');
        }
        
    } catch (e) {
        console.error('Save error:', e);
    }
}

/**
 * AUTO-LOAD INPUTS - Improved
 */
async function autoLoadInputs() {
    try {
        // Try localStorage first (faster)
        const saved = localStorage.getItem('refineryInputs');
        if (saved) {
            const savedInputs = JSON.parse(saved);
            applyInputValues(savedInputs);
            console.log('Inputs loaded from localStorage');
        }
        
        // Then try server (will override localStorage if successful)
        try {
            const response = await fetch(API_ENDPOINTS.LOAD_INPUTS);
            if (response.ok) {
                const serverInputs = await response.json();
                if (Object.keys(serverInputs).length > 0) {
                    applyInputValues(serverInputs);
                    console.log('Inputs loaded from server');
                }
            }
        } catch (serverError) {
            console.log('Server load failed, using localStorage');
        }
        
    } catch (e) {
        console.log('Load error:', e);
    }
}

/**
 * RUN SIMULATION
 */
async function runSimulation() {
    try {
        Utils.showLoading(true);

        const params = collectFormData();

        const response = await fetch(API_ENDPOINTS.SIMULATE, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(params)
        });

        if (!response.ok) {
            throw new Error('Simulation request failed');
        }

        currentResults = await response.json();

        if (currentResults.error) {
            alert('Simulation Error: ' + currentResults.error);
            return;
        }

        // Update solver recommended departure if in solver mode
        if (params.departureMode === 'solver' && currentResults.cargo_schedule && currentResults.cargo_schedule.length > 0) {
            document.getElementById('solverRecommendedDeparture').value = currentResults.cargo_schedule[0].dep_port;
        }

        // Display results
        displayResults(currentResults);
        displayInventoryTracking(currentResults.simulation_data); // Also display inventory tracking

        Utils.showResults();
        showTab('simulation', document.querySelector('.tab'));

    } catch (error) {
        console.error('Simulation error:', error);
        alert('Simulation failed: ' + error.message);
    } finally {
        Utils.showLoading(false);
    }
}

/**
 * DISPLAY RESULTS
 */
function displayResults(data) {
    // Display alerts
    const alertsContainer = document.getElementById('alertsContainer');
    alertsContainer.innerHTML = '<h3>⚠️ System Alerts</h3>';

    if (data.alerts && data.alerts.length > 0) {
        const alertsList = document.createElement('div');
        alertsList.className = 'alerts-list';

        data.alerts.forEach(alert => {
            const alertDiv = document.createElement('div');
            alertDiv.className = `alert alert-${alert.type}`;
            alertDiv.innerHTML = `<strong>Day ${alert.day}:</strong> ${alert.message}`;
            alertsList.appendChild(alertDiv);
        });

        alertsContainer.appendChild(alertsList);
    } else {
        alertsContainer.innerHTML += Utils.createAlert('success', '✅ No critical issues detected.');
    }

    // Display metrics
    const metricsContainer = document.getElementById('metricsContainer');
    metricsContainer.innerHTML = '<h3>📊 Performance Metrics</h3>';

    if (data.metrics) {
        const metricsDiv = document.createElement('div');
        metricsDiv.className = 'metrics-grid';
        // Safety checks added for metrics before calling .toFixed()
        const processingEfficiency = data.metrics.processing_efficiency ? data.metrics.processing_efficiency.toFixed(1) : 'N/A';
        const avgUtilization = data.metrics.avg_utilization ? data.metrics.avg_utilization.toFixed(1) : 'N/A';

        metricsDiv.innerHTML = `
            <div class="metric-card">
                <h4>Processing Efficiency</h4>
                <p class="metric-value">${processingEfficiency}%</p>
            </div>
            <div class="metric-card">
                <h4>Total Processed</h4>
                <p class="metric-value">${data.metrics.total_processed ? data.metrics.total_processed.toLocaleString() : 'N/A'} bbl</p>
            </div>
            <div class="metric-card">
                <h4>Critical Days</h4>
                <p class="metric-value">${data.metrics.critical_days} days</p>
            </div>
            <div class="metric-card">
                <h4>Tank Utilization</h4>
                <p class="metric-value">${avgUtilization}%</p>
            </div>
            <div class="metric-card">
                <h4>Clash Days</h4>
                <p class="metric-value">${data.metrics.clash_days} days</p>
            </div>
            <div class="metric-card">
                <h4>Sustainable</h4>
                <p class="metric-value">${data.metrics.sustainable_processing ? '✅ Yes' : '❌ No'}</p>
            </div>
        `;
        metricsContainer.appendChild(metricsDiv);
    }

    // Display cargo report
    displayCargoReport(data);

    // Display daily report
    displayDailyReport(data);
}

/**
 * DISPLAY DAILY REPORT
 */
function displayDailyReport(results) {
    const container = document.getElementById('dailyReportContainer');

    if (!results.simulation_data || results.simulation_data.length === 0) {
        container.innerHTML = '<p>No daily report data available</p>';
        return;
    }

    let tableHTML = `
        <h3>📊 Daily Operations Report</h3>
        <table class="schedule-table">
            <thead>
                <tr>
                    <th>Day</th>
                    <th>Date</th>
                    <th>Open Inventory</th>
                    <th>Processing</th>
                    <th>Closing Inventory</th>
                    <th>Tank Util %</th>
                    <th>Cargo Arrival</th>
                </tr>
            </thead>
            <tbody>
    `;

    results.simulation_data.forEach((dayData) => {
        const cargoInfo = dayData.cargo_type ? `${dayData.cargo_type} (${Utils.formatNumber(dayData.arrivals)})` : '-';
        // Added safety check for tank_utilization
        const tankUtilization = dayData.tank_utilization ? dayData.tank_utilization.toFixed(1) + '%' : 'N/A';
        tableHTML += `
            <tr>
                <td><strong>${dayData.day}</strong></td>
                <td>${dayData.date}</td>
                <td style="color: #007bff;">${Utils.formatNumber(dayData.start_inventory)}</td>
                <td style="color: #dc3545;">${Utils.formatNumber(dayData.processing)}</td>
                <td style="color: #28a745;">${Utils.formatNumber(dayData.end_inventory)}</td>
                <td style="color: #6f42c1;">${tankUtilization}</td>
                <td>${cargoInfo}</td>
            </tr>
        `;
    });

    tableHTML += '</tbody></table>';
    container.innerHTML = tableHTML;
}

/**
 * BUFFER ANALYSIS
 */
async function calculateBuffer() {
    try {
        Utils.showLoading(true);

        const params = collectFormData();

        const response = await fetch(API_ENDPOINTS.BUFFER_ANALYSIS, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(params)
        });

        if (!response.ok) {
            throw new Error('Buffer analysis request failed');
        }

        const bufferResults = await response.json();

        displayBufferAnalysis(bufferResults);
        Utils.showResults();
        showTab('buffer', document.querySelectorAll('.tab')[1]);

    } catch (error) {
        console.error('Buffer analysis error:', error);
        alert('Buffer analysis failed: ' + error.message);
    } finally {
        Utils.showLoading(false);
    }
}

/**
 * DISPLAY BUFFER ANALYSIS
 */
function displayBufferAnalysis(bufferResults) {
    const container = document.getElementById('bufferResults');

    let html = '<h3>🛡️ Buffer Analysis Report</h3>';

    if (bufferResults && Object.keys(bufferResults).length > 0) {
        html += '<div class="buffer-scenarios">';

        Object.entries(bufferResults).forEach(([scenarioKey, scenario]) => {
            const adequateText = scenario.adequate_current ? '✅ Adequate' : '❌ Insufficient';
            const adequateColor = scenario.adequate_current ? '#28a745' : '#dc3545';

            html += `
                <div class="scenario-card" style="border: 1px solid #ddd; margin: 10px 0; padding: 15px; border-radius: 5px;">
                    <h4>${scenario.description}</h4>
                    <div class="scenario-details">
                        <p><strong>Lead Time:</strong> ${scenario.lead_time.toFixed(1)} days</p>
                        <p><strong>Buffer Needed:</strong> ${Utils.formatNumber(scenario.buffer_needed)} barrels</p>
                        <p><strong>Tanks Required:</strong> ${scenario.tanks_needed} tanks</p>
                        <p><strong>Current Capacity:</strong> <span style="color: ${adequateColor}; font-weight: bold;">${adequateText}</span></p>
                        ${scenario.additional_tanks > 0 ?
                            `<p style="color: #dc3545;"><strong>Additional Tanks Needed:</strong> ${scenario.additional_tanks}</p>` :
                            '<p style="color: #28a745;"><strong>No additional tanks needed</strong></p>'
                        }
                    </div>
                </div>
            `;
        });

        html += '</div>';
    } else {
        html += '<p>No buffer analysis data available</p>';
    }

    container.innerHTML = html;
}

/**
 * CARGO OPTIMIZATION
 */
async function optimizeTanks() {
    try {
        Utils.showLoading(true);

        const params = collectFormData();

        const response = await fetch(API_ENDPOINTS.CARGO_OPTIMIZATION, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(params)
        });

        if (!response.ok) {
            throw new Error('Optimization request failed');
        }

        const optimizationResults = await response.json();

        displayCargoOptimizationResults(optimizationResults);
        Utils.showResults();
        showTab('optimization', document.querySelectorAll('.tab')[2]);

    } catch (error) {
        console.error('Cargo optimization error:', error);
        alert('Optimization failed: ' + error.message);
    } finally {
        Utils.showLoading(false);
    }
}

/**
 * DISPLAY CARGO OPTIMIZATION
 */
function displayCargoOptimizationResults(optimizationResults) {
    const container = document.getElementById('optimizationResults');

    let html = '<h3>⚡ Cargo Optimization Report</h3>';

    if (optimizationResults && Object.keys(optimizationResults).length > 0) {
        html += '<div class="optimization-combos">';

        Object.entries(optimizationResults).forEach(([comboKey, combo]) => {
            const sustainableText = combo.sustainable ? '✅ Sustainable' : '❌ Not Sustainable';
            const sustainableColor = combo.sustainable ? '#28a745' : '#dc3545';

            html += `
                <div class="combo-card" style="border: 1px solid #ddd; margin: 10px 0; padding: 15px; border-radius: 5px;">
                    <h4>Combination ${comboKey.replace('combo_', '')}: ${combo.cargo_types.join(' + ').toUpperCase()}</h4>
                    <div class="combo-details">
                        <p><strong>Processing Efficiency:</strong> ${combo.efficiency.toFixed(1)}%</p>
                        <p><strong>Total Cargoes:</strong> ${combo.total_cargoes}</p>
                        <p><strong>Cargo Mix:</strong> ${combo.cargo_mix}</p>
                        <p><strong>Clash Days:</strong> ${combo.clash_days}</p>
                        <p><strong>Min Inventory:</strong> ${Utils.formatNumber(combo.min_inventory)} bbl</p>
                        <p><strong>Operations:</strong> <span style="color: ${sustainableColor}; font-weight: bold;">${sustainableText}</span></p>
                    </div>
                </div>
            `;
        });

        html += '</div>';

        // Find best combination
        const bestCombo = Object.values(optimizationResults).reduce((best, current) => {
            if (current.sustainable && current.efficiency > (best?.efficiency || 0)) {
                return current;
            }
            return best;
        }, null);

        if (bestCombo) {
            html += `
                <div style="background-color: #e7f5e7; padding: 15px; border-radius: 5px; margin-top: 15px;">
                    <h4 style="color: #28a745;">💡 Recommended: ${bestCombo.cargo_types.join(' + ').toUpperCase()}</h4>
                    <p>Best efficiency (${bestCombo.efficiency.toFixed(1)}%) with sustainable operations</p>
                </div>
            `;
        }
    } else {
        html += '<p>No optimization data available</p>';
    }

    container.innerHTML = html;
}

/**
 * SHOW TANK STATUS
 */
async function showTankStatus() {
    if (!currentResults) {
        alert('Please run a simulation first');
        return;
    }

    try {
        Utils.showLoading(true);

        const response = await fetch(API_ENDPOINTS.EXPORT_TANK_STATUS, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(currentResults)
        });

        if (!response.ok) {
            throw new Error('Tank status export failed');
        }

        const result = await response.json();
        alert(`✅ Tank status exported: ${result.filename}`);

    } catch (error) {
        console.error('Tank status error:', error);
        alert('Tank status export failed: ' + error.message);
    } finally {
        Utils.showLoading(false);
    }
}

/**
 * EXPORT SIMULATION REPORT
 */
async function exportSimulationReport() {
    try {
        Utils.showLoading(true);

        if (!currentResults) {
            alert('Please run a simulation first before exporting.');
            Utils.showLoading(false);
            return;
        }

        const response = await fetch(API_ENDPOINTS.EXPORT_TANK_STATUS, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(currentResults)
        });

        if (!response.ok) {
            throw new Error('Export failed');
        }

        const result = await response.json();
        alert(`✅ Simulation report exported: ${result.filename}`);

    } catch (error) {
        console.error('Export error:', error);
        alert('Export failed: ' + error.message);
    } finally {
        Utils.showLoading(false);
    }
}

/**
 * DISPLAY CARGO REPORT
 */
function displayCargoReport(data) {
    const container = document.getElementById('cargoReportContainer');
    if (!container) return;

    if (!data.cargo_report || data.cargo_report.length === 0) {
        container.innerHTML = '<h3>🚢 Cargo Schedule Report</h3><p><em>No cargo schedule available</em></p>';
        return;
    }

    const cargoReport = data.cargo_report;

    let html = '<h3>🚢 Cargo Schedule Report</h3>';
    html += '<p><em>Detailed cargo timeline with load port, departure, arrival, and discharge times</em></p>';
    html += '<div class="cargo-schedule-table">';
    html += '<table class="data-table">';
    html += '<thead><tr>';
    html += '<th>Cargo Type</th>';
    html += '<th>Load Port Time</th>';
    html += '<th>Dep Time</th>';
    html += '<th>Arrival Time</th>';
    html += '<th>Dep Unload Port</th>';
    html += '<th>Cargo Size</th>';
    html += '</tr></thead><tbody>';

    cargoReport.forEach(cargo => {
        html += '<tr>';
        html += `<td>${cargo.Cargo_type}</td>`;
        html += `<td>${cargo.Load_Port_time}</td>`;
        html += `<td>${cargo.dep_time}</td>`;
        html += `<td>${cargo.Arrival_time}</td>`;
        html += `<td>${cargo.dep_unload_port}</td>`;
        html += `<td>${cargo.Cargo_size}</td>`;
        html += '</tr>';
    });

    html += '</tbody></table></div>';

    container.innerHTML = html;
}

/**
 * SHOW TAB
 */
function showTab(tabId, tabButton) {
    // Hide all tab contents
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });

    // Remove active class from all tabs
    document.querySelectorAll('.tab').forEach(tab => {
        tab.classList.remove('active');
    });

    // Show selected tab content and activate button
    document.getElementById(tabId).classList.add('active');
    if (tabButton) tabButton.classList.add('active');
}

/**
 * Validate inventory range inputs in real-time - Updated for dynamic tanks
 */
function validateInventoryRange() {
    const minInventory = parseFloat(document.getElementById('minInventory').value) || 0;
    const maxInventory = parseFloat(document.getElementById('maxInventory').value) || 0;
    const messageDiv = document.getElementById('inventoryValidationMessage');
    // Count actual tank boxes in DOM instead of input field
    const actualTankCount = document.querySelectorAll('.tank-box').length;

    let isValid = true;
    let message = '';
    let messageType = 'success';

    if (minInventory >= maxInventory) {
        isValid = false;
        message = '❌ Minimum inventory must be less than maximum inventory';
        messageType = 'error';
    } else if (minInventory < 0 || maxInventory < 0) {
        isValid = false;
        message = '❌ Inventory values cannot be negative';
        messageType = 'error';
    } else {
        // Calculate current inventory for all actual tanks
        let currentInventory = 0;
        const tankLevelInputs = document.querySelectorAll('input[id*="Level"]');
        tankLevelInputs.forEach(input => {
            if (input.id.includes('tank') && input.id.includes('Level')) {
                const tankNumber = input.id.replace('tank', '').replace('Level', '');
                const tankLevel = parseFloat(input.value) || 0;
                const deadBottom = parseFloat(document.getElementById(`deadBottom${tankNumber}`)?.value) || 10000;
                currentInventory += Math.max(0, tankLevel - deadBottom);
            }
        });

        if (currentInventory < minInventory) {
            isValid = false;
            message = `⚠️ Current inventory (${currentInventory.toLocaleString()} bbl) is below minimum (${minInventory.toLocaleString()} bbl)`;
            messageType = 'warning';
        } else if (currentInventory > maxInventory) {
            isValid = false;
            message = `⚠️ Current inventory (${currentInventory.toLocaleString()} bbl) is above maximum (${maxInventory.toLocaleString()} bbl)`;
            messageType = 'warning';
        } else {
            message = `✅ Current inventory: ${currentInventory.toLocaleString()} bbl (Range: ${minInventory.toLocaleString()} - ${maxInventory.toLocaleString()} bbl) - ${actualTankCount} tanks`;
            messageType = 'success';
        }
    }

    // Display message
    if (messageDiv) {
        messageDiv.style.display = 'block';
        messageDiv.innerHTML = message;

        if (messageType === 'error') {
            messageDiv.style.backgroundColor = '#f8d7da';
            messageDiv.style.color = '#721c24';
            messageDiv.style.border = '1px solid #f5c6cb';
        } else if (messageType === 'warning') {
            messageDiv.style.backgroundColor = '#fff3cd';
            messageDiv.style.color = '#856404';
            messageDiv.style.border = '1px solid #ffeaa7';
        } else {
            messageDiv.style.backgroundColor = '#d1edff';
            messageDiv.style.color = '#0c5460';
            messageDiv.style.border = '1px solid #bee5eb';
        }
    }

    return isValid;
}

/**
 * INVENTORY button click handler
 */
function checkInventoryRange() {
    Utils.showLoading(true);

    const params = collectFormData();

    fetch('/api/validate_inventory_range', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(params)
        })
        .then(response => response.json())
        .then(data => {
            Utils.showLoading(false);

            if (data.success) {
                alert(`✅ INVENTORY RANGE VALIDATION PASSED\n\n${data.message}\n\nYou can proceed with simulation.`);
            } else {
                alert(`❌ INVENTORY RANGE VALIDATION FAILED\n\n${data.message}\n\nPlease adjust your inventory range or tank levels.`);
            }
        })
        .catch(error => {
            Utils.showLoading(false);
            console.error('Inventory validation error:', error);
            alert('❌ Error validating inventory range. Please try again.');
        });
}

/**
 * Display inventory tracking results
 */
function displayInventoryTracking(inventoryData) {
    const container = document.getElementById('inventoryResults');
    if (!container || !inventoryData || inventoryData.length === 0) {
        if (container) container.innerHTML = '<p>No inventory tracking data available.</p>';
        return;
    }
    
    // Setup for Chart.js
    const ctx = document.getElementById('inventoryChart').getContext('2d');
    const labels = inventoryData.map(d => `Day ${d.day}`);
    const dataPoints = inventoryData.map(d => d.end_inventory);

    // Destroy existing chart if it exists to prevent conflicts
    if (window.myInventoryChart) {
        window.myInventoryChart.destroy();
    }

    // Create the chart
    window.myInventoryChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'End of Day Inventory (bbl)',
                data: dataPoints,
                borderColor: '#007bff',
                backgroundColor: 'rgba(0, 123, 255, 0.1)',
                fill: true,
                tension: 0.1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value, index, values) {
                            return value.toLocaleString() + ' bbl';
                        }
                    }
                }
            }
        }
    });
}

/**
 * Enhanced runSimulation function with inventory validation
 */
function runSimulationWithInventoryCheck() {
    // First validate inventory range
    const minInventory = parseFloat(document.getElementById('minInventory').value) || 0;
    const maxInventory = parseFloat(document.getElementById('maxInventory').value) || 0;

    if (minInventory > 0 || maxInventory > 0) {
        if (minInventory >= maxInventory) {
            alert('❌ SIMULATION BLOCKED\n\nMinimum inventory must be less than maximum inventory.\nPlease fix inventory range before running simulation.');
            return;
        }
    }

    // Proceed with normal simulation
    runSimulation();
}

/**
 * Update collectSimulationParams to include inventory range
 */
function collectSimulationParamsWithInventory() {
    const params = collectFormData();

    // Add inventory range parameters
    params.minInventory = parseFloat(document.getElementById('minInventory').value) || 0;
    params.maxInventory = parseFloat(document.getElementById('maxInventory').value) || 0;

    return params;
}

// CORRECT - handles file download


/**
 * Enhanced export function to handle inventory data
 */
function exportSimulationReportWithInventory() {
    if (!currentResults) {
        alert('Please run a simulation first before exporting.');
        return;
    }

    Utils.showLoading(true);

    fetch('/api/export_tank_status', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(currentResults)
        })
        .then(response => response.json())
        .then(data => {
            Utils.showLoading(false);
            if (data.success) {
                alert(`✅ COMPLETE EXPORT SUCCESSFUL\n\nFile: ${data.filename}\n\nFeatures included:\n${data.features?.join('\n') || 'All fixed requirements implemented'}\n\nIncluding INVENTORY sheet with real-time graph!`);
            } else {
                alert(`❌ Export failed: ${data.error}`);
            }
        })
        .catch(error => {
            Utils.showLoading(false);
            console.error('Export error:', error);
            alert('❌ Export failed. Please try again.');
        });
}

function initializeAutoSave() {
    // Get all input and select elements
    const inputs = document.querySelectorAll('input, select');
    
    inputs.forEach(input => {
        // Add event listeners for different input types
        if (input.type === 'number' || input.type === 'text') {
            // For text/number inputs, save on blur (when user finishes editing)
            input.addEventListener('blur', autoSaveInputs);
            // Also save on input change with debouncing
            let timeout;
            input.addEventListener('input', () => {
                clearTimeout(timeout);
                timeout = setTimeout(autoSaveInputs, 1000); // Save after 1 second of no typing
            });
        } else {
            // For select, radio, checkbox - save immediately on change
            input.addEventListener('change', autoSaveInputs);
        }
    });
    
    console.log(`Auto-save initialized for ${inputs.length} inputs`);
}





// ADD SAVE STATUS INDICATOR (Optional visual feedback)
function showSaveStatus(status) {
    // Create or update save status indicator
    let indicator = document.getElementById('saveIndicator');
    if (!indicator) {
        indicator = document.createElement('div');
        indicator.id = 'saveIndicator';
        indicator.style.cssText = `
            position: fixed;
            top: 10px;
            right: 10px;
            padding: 8px 12px;
            background: #28a745;
            color: white;
            border-radius: 4px;
            font-size: 12px;
            z-index: 1000;
            transition: opacity 0.3s;
        `;
        document.body.appendChild(indicator);
    }
    
    if (status === 'saved') {
        indicator.textContent = '✓ Saved';
        indicator.style.opacity = '1';
        setTimeout(() => {
            indicator.style.opacity = '0';
        }, 2000);
    }
}

// HELPER FUNCTION TO APPLY VALUES
function applyInputValues(inputValues) {
    Object.entries(inputValues).forEach(([id, value]) => {
        const element = document.getElementById(id);
        if (element) {
            if (element.type === 'checkbox') {
                element.checked = value;
            } else {
                element.value = value;
            }
        }
    });

    // Update tank count if it was saved
    if (inputValues.numTanks) {
        updateTankCount();
    }

    // Update calculations after loading
    toggleDepartureMode();
    autoCalculateLeadTime();
    autoCalculatePumpingDays();
    validateInventoryRange();
}

async function exportCharts() {
    // Check if simulation has been run
    if (!currentResults) {
        alert('⚠️ Please run a simulation first to generate charts data.');
        return;
    }

    try {
        // Show loading spinner
        Utils.showLoading(true);
        document.getElementById('loading').querySelector('p').textContent = 'Generating charts...';
        
        // Send simulation results to backend for chart generation
        const response = await fetch('/api/export_charts', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(currentResults)
        });

        const result = await response.json();

        if (result.success) {
            alert(`✅ ${result.message}`);
        } else {
            alert(`❌ Charts export failed: ${result.error}`);
        }
        
    } catch (error) {
        console.error('Charts export error:', error);
        alert(`❌ Charts export error: ${error.message}`);
    } finally {
        // Hide loading spinner
        Utils.showLoading(false);
        document.getElementById('loading').querySelector('p').textContent = 'Running simulation...';
    }
}

function scrollToReport() {
    const element = document.getElementById('dailyReportContainer');
    if (element) {
        element.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
}

// CONSOLIDATED INITIALIZATION - Merged from HTML and existing main.js
document.addEventListener('DOMContentLoaded', () => {
    // Load saved inputs first
    autoLoadInputs();
    
    // Initialize calculations (from HTML)
    autoCalculateLeadTime();
    autoCalculatePumpingDays();
    validateInventoryRange();
    
    // Update tank count to create missing tanks
    setTimeout(() => {
        updateTankCount();
        initializeAutoSave();
    }, 500); // Small delay to ensure all elements are loaded
});


/**
 * FIXED: Show Tank Status - Handle file download properly
 */
async function showTankStatus() {
    if (!currentResults) {
        alert('Please run a simulation first');
        return;
    }

    try {
        Utils.showLoading(true);

        const response = await fetch(API_ENDPOINTS.EXPORT_TANK_STATUS, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(currentResults)
        });

        if (!response.ok) {
            throw new Error('Tank status export failed');
        }

        // Check if response is a file or JSON
        const contentType = response.headers.get('content-type');
        
        if (contentType && contentType.includes('application/json')) {
            // It's JSON - parse it
            const result = await response.json();
            alert(`✅ Tank status exported: ${result.filename}`);
        } else {
            // It's a file - trigger download
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            
            // Get filename from Content-Disposition header or use default
            const disposition = response.headers.get('Content-Disposition');
            let filename = 'tank_status_export.xlsx';
            if (disposition) {
                const matches = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disposition);
                if (matches != null && matches[1]) {
                    filename = matches[1].replace(/['"]/g, '');
                }
            }
            
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
            alert(`✅ Tank status downloaded: ${filename}`);
        }

    } catch (error) {
        console.error('Tank status error:', error);
        alert('Tank status export failed: ' + error.message);
    } finally {
        Utils.showLoading(false);
    }
}

/**
 * FIXED: Export Simulation Report - Handle file download properly
 */
async function exportSimulationReport() {
    try {
        Utils.showLoading(true);

        if (!currentResults) {
            alert('Please run a simulation first before exporting.');
            Utils.showLoading(false);
            return;
        }

        const response = await fetch(API_ENDPOINTS.EXPORT_TANK_STATUS, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(currentResults)
        });

        if (!response.ok) {
            throw new Error('Export failed');
        }

        // Check if response is a file or JSON
        const contentType = response.headers.get('content-type');
        
        if (contentType && contentType.includes('application/json')) {
            // It's JSON - parse it
            const result = await response.json();
            alert(`✅ Simulation report exported: ${result.filename}`);
        } else {
            // It's a file - trigger download
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            
            // Get filename from Content-Disposition header or use default
            const disposition = response.headers.get('Content-Disposition');
            let filename = 'simulation_report.xlsx';
            if (disposition) {
                const matches = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disposition);
                if (matches != null && matches[1]) {
                    filename = matches[1].replace(/['"]/g, '');
                }
            }
            
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
            alert(`✅ Simulation report downloaded: ${filename}`);
        }

    } catch (error) {
        console.error('Export error:', error);
        alert('Export failed: ' + error.message);
    } finally {
        Utils.showLoading(false);
    }
}

/**
 * FIXED: Export Simulation Report With Inventory - Handle file download properly
 */
function exportSimulationReportWithInventory() {
    if (!currentResults) {
        alert('Please run a simulation first before exporting.');
        return;
    }

    Utils.showLoading(true);

    fetch('/api/export_tank_status', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(currentResults)
        })
        .then(async response => {
            if (!response.ok) {
                throw new Error('Export failed');
            }

            // Check if response is a file or JSON
            const contentType = response.headers.get('content-type');
            
            if (contentType && contentType.includes('application/json')) {
                // It's JSON - parse it
                const data = await response.json();
                Utils.showLoading(false);
                if (data.success) {
                    alert(`✅ COMPLETE EXPORT SUCCESSFUL\n\nFile: ${data.filename}\n\nFeatures included:\n${data.features?.join('\n') || 'All fixed requirements implemented'}\n\nIncluding INVENTORY sheet with real-time graph!`);
                } else {
                    alert(`❌ Export failed: ${data.error}`);
                }
            } else {
                // It's a file - trigger download
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                
                // Get filename from Content-Disposition header or use default
                const disposition = response.headers.get('Content-Disposition');
                let filename = 'complete_simulation_report.xlsx';
                if (disposition) {
                    const matches = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disposition);
                    if (matches != null && matches[1]) {
                        filename = matches[1].replace(/['"]/g, '');
                    }
                }
                
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
                
                Utils.showLoading(false);
                alert(`✅ COMPLETE EXPORT SUCCESSFUL\n\nFile downloaded: ${filename}\n\nIncluding INVENTORY sheet with real-time graph!`);
            }
        })
        .catch(error => {
            Utils.showLoading(false);
            console.error('Export error:', error);
            alert('❌ Export failed. Please try again.');
        });
}

/**
 * FIXED: Export Charts - Handle file download properly
 */
async function exportCharts() {
    // Check if simulation has been run
    if (!currentResults) {
        alert('⚠️ Please run a simulation first to generate charts data.');
        return;
    }

    try {
        // Show loading spinner
        Utils.showLoading(true);
        document.getElementById('loading').querySelector('p').textContent = 'Generating charts...';
        
        // Send simulation results to backend for chart generation
        const response = await fetch('/api/export_charts', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(currentResults)
        });

        if (!response.ok) {
            throw new Error('Charts export failed');
        }

        // Check if response is a file or JSON
        const contentType = response.headers.get('content-type');
        
        if (contentType && contentType.includes('application/json')) {
            // It's JSON - parse it
            const result = await response.json();
            if (result.success) {
                alert(`✅ ${result.message}`);
            } else {
                alert(`❌ Charts export failed: ${result.error}`);
            }
        } else {
            // It's a file - trigger download
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            
            // Get filename from Content-Disposition header or use default
            const disposition = response.headers.get('Content-Disposition');
            let filename = 'charts_export.xlsx';
            if (disposition) {
                const matches = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disposition);
                if (matches != null && matches[1]) {
                    filename = matches[1].replace(/['"]/g, '');
                }
            }
            
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
            alert(`✅ Charts exported and downloaded: ${filename}`);
        }
        
    } catch (error) {
        console.error('Charts export error:', error);
        alert(`❌ Charts export error: ${error.message}`);
    } finally {
        // Hide loading spinner
        Utils.showLoading(false);
        document.getElementById('loading').querySelector('p').textContent = 'Running simulation...';
    }
}


// Make all functions globally available
window.populateTankLevels = populateTankLevels;
window.autoCalculatePumpingDays = autoCalculatePumpingDays;
window.autoCalculateLeadTime = autoCalculateLeadTime;
window.toggleDepartureMode = toggleDepartureMode;
window.applyDefaultDeadBottom = applyDefaultDeadBottom;
window.autoSaveInputs = autoSaveInputs;
window.autoLoadInputs = autoLoadInputs;
window.runSimulation = runSimulation;
window.calculateBuffer = calculateBuffer;
window.optimizeTanks = optimizeTanks;
window.showTankStatus = showTankStatus;
window.exportSimulationReport = exportSimulationReport;
window.showTab = showTab;
window.validateInventoryRange = validateInventoryRange;
window.checkInventoryRange = checkInventoryRange;
window.runSimulationWithInventoryCheck = runSimulationWithInventoryCheck;
window.exportSimulationReportWithInventory = exportSimulationReportWithInventory;

// Add the moved functions to global window object
window.scrollToTop = scrollToTop;
window.scrollToCargoReport = scrollToCargoReport;
window.scrollToBottom = scrollToBottom;
window.scrollToSimulation = scrollToSimulation;
window.updateTankCount = updateTankCount;
window.addOneTank = addOneTank;
window.addNewTankBox = addNewTankBox;
window.initializeAutoSave = initializeAutoSave;
window.showSaveStatus = showSaveStatus;
window.applyInputValues = applyInputValues;
window.getCurrentTankCount = getCurrentTankCount;
window.exportCharts = exportCharts;