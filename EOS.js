// Global variables
let appData = JSON.parse(localStorage.getItem('energyOptimizationData')) || [];
let costChart = null;
let suggestionInterval = null;

// DOM elements
const interfaces = {
    welcome: document.getElementById('welcome-interface'),
    input: document.getElementById('input-interface'),
    result: document.getElementById('result-interface'),
    dataAccess: document.getElementById('data-access-interface'),
    dataView: document.getElementById('data-view-interface')
};

const buttons = {
    getStarted: document.getElementById('get-started-btn'),
    calculate: document.getElementById('calculate-btn'),
    viewData: document.getElementById('view-data-btn'),
    saveData: document.getElementById('save-data-btn'),
    newCalculation: document.getElementById('new-calculation-btn'),
    goToData: document.getElementById('go-to-data-btn'),
    submitPassword: document.getElementById('submit-password-btn'),
    backToInput: document.getElementById('back-to-input-btn'),
    backToInputFromData: document.getElementById('back-to-input-from-data-btn'),
    exportData: document.getElementById('export-data-btn')
};

const sound = document.getElementById('click-sound');

// Appliance data
const appliances = [
    { name: 'Fan', watts: 70, icon: '🎐' },
    { name: 'Tube Light', watts: 40, icon: '💡' },
    { name: 'LED Bulb', watts: 9, icon: '💡' },
    { name: 'Refrigerator', watts: 150, icon: '❄️' },
    { name: 'TV', watts: 100, icon: '📺' },
    { name: 'Washing Machine', watts: 500, icon: '🧺' },
    { name: 'Water Pump', watts: 750, icon: '🚿' },
    { name: 'Geyser', watts: 2000, icon: '♨️' },
    { name: 'Mixer', watts: 500, icon: '🍲' },
    { name: 'Air Conditioner', watts: 1000, icon: '❄️' },
    { name: 'Iron', watts: 1000, icon: '👔' }
];

// Renewable energy suggestions
const renewableSuggestions = {
    solar: [
        "Install high-efficiency monocrystalline solar panels",
        "Consider solar water heating for your geyser",
        "Explore grid-tied systems with net metering",
        "South-facing roofs get the most sunlight for solar panels",
        "Solar panels typically last 25-30 years with minimal maintenance"
    ],
    wind: [
        "Small wind turbines can be effective in windy areas",
        "Wind energy is most consistent in hilly and coastal regions",
        "Consider a hybrid system with both wind and solar for reliability",
        "Wind turbines need to be placed at least 30 feet above obstacles",
        "Check local regulations for wind turbine installations"
    ],
    hydro: [
        "Micro-hydro systems are great for properties with flowing water",
        "Rainwater harvesting can supplement your water source",
        "Pico-hydro systems can work with small streams",
        "Hydro systems provide consistent 24/7 power generation",
        "Requires proper water rights and permits"
    ],
    hybrid: [
        "A hybrid system combines solar and wind for more consistent energy",
        "Battery storage can help store excess energy for later use",
        "Hybrid systems are great for areas with unpredictable weather",
        "Consider a grid-tied system with net metering to sell excess power",
        "Hybrid systems can provide power even when one source is unavailable"
    ],
    general: [
        "Switch to energy-efficient appliances to reduce consumption",
        "Use natural lighting during the day to reduce electricity usage",
        "Unplug devices when not in use to prevent phantom loads",
        "Regular maintenance of appliances improves their efficiency",
        "Consider energy audits to identify more savings opportunities"
    ]
};

// Event Listeners
document.addEventListener('DOMContentLoaded', () => {
    // Initialize appliance list
    renderApplianceList();
    
    // Set up event listeners
    buttons.getStarted.addEventListener('click', () => switchInterface('input'));
    buttons.calculate.addEventListener('click', calculateResults);
    buttons.viewData.addEventListener('click', () => switchInterface('dataAccess'));
    buttons.saveData.addEventListener('click', saveCurrentData);
    buttons.newCalculation.addEventListener('click', () => {
        clearInterval(suggestionInterval);
        switchInterface('input');
    });
    buttons.goToData.addEventListener('click', () => switchInterface('dataAccess'));
    buttons.submitPassword.addEventListener('click', checkPassword);
    buttons.backToInput.addEventListener('click', () => switchInterface('input'));
    buttons.backToInputFromData.addEventListener('click', () => switchInterface('input'));
    buttons.exportData.addEventListener('click', exportToExcel);
    
    // Set up other event listeners
    document.getElementById('location').addEventListener('change', updateCostOptions);
    document.getElementById('cost-per-kwh').addEventListener('change', toggleCustomCost);
});

// Functions
function playClickSound() {
    sound.currentTime = 0;
    sound.play();
}

function switchInterface(target) {
    playClickSound();
    
    // Hide all interfaces
    Object.values(interfaces).forEach(int => {
        int.classList.remove('active');
    });
    
    // Show target interface
    interfaces[target].classList.add('active');
    interfaces[target].classList.add('fade-in');
    
    // Special cases
    if (target === 'dataView') {
        displaySavedData();
    }
    
    // Scroll to top
    window.scrollTo(0, 0);
}

function renderApplianceList() {
    const container = document.querySelector('.appliance-list');
    container.innerHTML = '';
    
    appliances.forEach(appliance => {
        const item = document.createElement('div');
        item.className = 'appliance-item';
        item.innerHTML = `
            <h4>${appliance.icon} ${appliance.name} (${appliance.watts}W)</h4>
            <div class="appliance-input">
                <label>Qty:</label>
                <input type="number" class="appliance-qty" data-watts="${appliance.watts}" min="0" placeholder="0">
            </div>
            <div class="appliance-input">
                <label>Hours:</label>
                <input type="number" class="appliance-hours" min="0" max="24" placeholder="0">
            </div>
        `;
        container.appendChild(item);
    });
    
    // Add other appliances option
    const otherItem = document.createElement('div');
    otherItem.className = 'appliance-item';
    otherItem.innerHTML = `
        <h4>➕ Other Appliances</h4>
        <div class="appliance-input">
            <select class="other-appliance-type">
                <option value="">Select</option>
                <option value="5">Charger (5W)</option>
                <option value="10">Charger (10W)</option>
                <option value="15">Charger (15W)</option>
                <option value="20">Charger (20W)</option>
                <option value="2">Torch (2W)</option>
                <option value="5">Torch (5W)</option>
                <option value="10">Torch (10W)</option>
                <option value="10">Lamp (10W)</option>
                <option value="20">Lamp (20W)</option>
                <option value="30">Lamp (30W)</option>
                <option value="40">Lamp (40W)</option>
                <option value="50">Lamp (50W)</option>
                <option value="60">Lamp (60W)</option>
                <option value="custom">Custom Wattage</option>
            </select>
        </div>
        <div class="appliance-input custom-watts-container" style="display: none;">
            <input type="number" class="custom-watts" min="1" placeholder="Watts">
        </div>
        <div class="appliance-input">
            <input type="number" class="other-appliance-qty" min="0" placeholder="Qty">
        </div>
        <div class="appliance-input">
            <input type="number" class="other-appliance-hours" min="0" max="24" placeholder="Hours">
        </div>
    `;
    container.appendChild(otherItem);
    
    // Add event listener for custom watts
    container.querySelector('.other-appliance-type').addEventListener('change', function() {
        const customContainer = this.closest('.appliance-item').querySelector('.custom-watts-container');
        customContainer.style.display = this.value === 'custom' ? 'flex' : 'none';
    });
}

function updateCostOptions() {
    const location = document.getElementById('location').value;
    const costSelect = document.getElementById('cost-per-kwh');
    
    // Clear existing options except the first and custom
    while (costSelect.options.length > 1) {
        if (costSelect.options[1].value !== 'custom') {
            costSelect.remove(1);
        } else {
            break;
        }
    }
    
    // Add options based on location
    if (location === 'urban') {
        addCostOption('₹6 (Urban)', '6');
        addCostOption('₹7 (Urban)', '7');
        addCostOption('₹8 (Urban)', '8');
    } else if (location === 'rural') {
        addCostOption('₹4 (Rural)', '4');
        addCostOption('₹5 (Rural)', '5');
        addCostOption('₹6 (Rural)', '6');
    } else if (location === 'hilly') {
        addCostOption('₹5 (Hilly)', '5');
        addCostOption('₹6 (Hilly)', '6');
        addCostOption('₹7 (Hilly)', '7');
    } else if (location === 'coastal') {
        addCostOption('₹7 (Coastal)', '7');
        addCostOption('₹8 (Coastal)', '8');
        addCostOption('₹9 (Coastal)', '9');
    } else {
        // For semi-urban or no location selected, show all options
        addCostOption('₹6 (Urban)', '6');
        addCostOption('₹7 (Urban)', '7');
        addCostOption('₹8 (Urban)', '8');
        addCostOption('₹4 (Rural)', '4');
        addCostOption('₹5 (Rural/Hilly)', '5');
        addCostOption('₹6 (Rural)', '6');
        addCostOption('₹7 (Hilly/Coastal)', '7');
        addCostOption('₹8 (Coastal)', '8');
        addCostOption('₹9 (Coastal)', '9');
    }
    
    // Add custom option if not already present
    if (!Array.from(costSelect.options).some(opt => opt.value === 'custom')) {
        addCostOption('Custom', 'custom');
    }
}

function addCostOption(text, value) {
    const option = document.createElement('option');
    option.text = text;
    option.value = value;
    document.getElementById('cost-per-kwh').add(option);
}

function toggleCustomCost() {
    const costSelect = document.getElementById('cost-per-kwh');
    const customCostInput = document.getElementById('custom-cost');
    
    if (costSelect.value === 'custom') {
        customCostInput.style.display = 'block';
        customCostInput.required = true;
    } else {
        customCostInput.style.display = 'none';
        customCostInput.required = false;
    }
}

function calculateResults() {
    playClickSound();
    
    // Validate required fields
    const houseNumber = document.getElementById('house-number').value;
    const meterNumber = document.getElementById('meter-number').value;
    const location = document.getElementById('location').value;
    const weather = document.getElementById('weather').value;
    const costSelect = document.getElementById('cost-per-kwh');
    const costValue = costSelect.value === 'custom' ? 
        parseFloat(document.getElementById('custom-cost').value) : 
        parseFloat(costSelect.value);
    
    if (!houseNumber || !meterNumber || !location || !weather || isNaN(costValue) || costValue <= 0) {
        alert('Please fill in all required fields with valid values.');
        return;
    }
    
    // Calculate appliance usage
    let totalWattsPerDay = 0;
    const applianceDetails = [];
    
    // Standard appliances
    document.querySelectorAll('.appliance-item:not(:last-child)').forEach((item, index) => {
        const qty = parseFloat(item.querySelector('.appliance-qty').value) || 0;
        const hours = parseFloat(item.querySelector('.appliance-hours').value) || 0;
        const watts = appliances[index].watts;
        const dailyUsage = qty * watts * hours;
        const monthlyUsage = dailyUsage * 30 / 1000; // Convert to kWh
        
        if (qty > 0 && hours > 0) {
            totalWattsPerDay += dailyUsage;
            applianceDetails.push({
                name: appliances[index].name,
                qty,
                hours,
                monthlyUsage: parseFloat(monthlyUsage.toFixed(2))
            });
        }
    });
    
    // Other appliances
    const otherItem = document.querySelector('.appliance-item:last-child');
    const otherTypeSelect = otherItem.querySelector('.other-appliance-type');
    const otherTypeValue = otherTypeSelect.value;
    let otherWatts = 0;
    
    if (otherTypeValue && otherTypeValue !== '') {
        if (otherTypeValue === 'custom') {
            otherWatts = parseFloat(otherItem.querySelector('.custom-watts').value) || 0;
        } else {
            otherWatts = parseFloat(otherTypeValue);
        }
        
        const otherQty = parseFloat(otherItem.querySelector('.other-appliance-qty').value) || 0;
        const otherHours = parseFloat(otherItem.querySelector('.other-appliance-hours').value) || 0;
        const dailyUsage = otherQty * otherWatts * otherHours;
        const monthlyUsage = dailyUsage * 30 / 1000; // Convert to kWh
        
        if (otherQty > 0 && otherHours > 0) {
            totalWattsPerDay += dailyUsage;
            applianceDetails.push({
                name: 'Other Appliance',
                qty: otherQty,
                hours: otherHours,
                monthlyUsage: parseFloat(monthlyUsage.toFixed(2))
            });
        }
    }
    
    // Calculate total consumption and bill
    const totalConsumption = totalWattsPerDay * 30 / 1000; // kWh/month
    const totalBill = totalConsumption * costValue;
    
    // Determine renewable energy recommendation
    const { 
        renewableSource, 
        installationCost, 
        reductionPercentage, 
        suggestions,
        benefits,
        considerations
    } = getRenewableRecommendation(location, weather, totalConsumption);
    
    // Calculate new bill and savings
    const newBill = totalBill * (1 - reductionPercentage / 100);
    const monthlySavings = totalBill - newBill;
    const yearlySavings = monthlySavings * 12;
    
    // Display results
    document.getElementById('total-consumption').textContent = `${totalConsumption.toFixed(2)} kWh/month`;
    document.getElementById('total-bill').textContent = `₹${totalBill.toFixed(2)}`;
    document.getElementById('renewable-source').textContent = renewableSource;
    document.getElementById('installation-cost').textContent = `₹${installationCost.toLocaleString()}`;
    document.getElementById('new-bill').textContent = `₹${newBill.toFixed(2)}`;
    document.getElementById('monthly-savings').textContent = `₹${monthlySavings.toFixed(2)}`;
    document.getElementById('yearly-savings').textContent = `₹${yearlySavings.toFixed(2)}`;
    
    // Display renewable energy details
    displayRenewableDetails(renewableSource, benefits, considerations);
    
    // Display rotating suggestions
    displayRotatingSuggestions(suggestions);
    
    // Display comparison table
    const comparisonData = [
        { metric: 'Monthly Cost', before: `₹${totalBill.toFixed(2)}`, after: `₹${newBill.toFixed(2)}`, savings: `₹${monthlySavings.toFixed(2)}` },
        { metric: 'Yearly Cost', before: `₹${(totalBill * 12).toFixed(2)}`, after: `₹${(newBill * 12).toFixed(2)}`, savings: `₹${yearlySavings.toFixed(2)}` },
        { metric: 'Carbon Footprint', before: `${(totalConsumption * 0.85).toFixed(2)} kg CO2`, after: `${(totalConsumption * 0.85 * (1 - reductionPercentage / 100)).toFixed(2)} kg CO2`, savings: `${(totalConsumption * 0.85 * reductionPercentage / 100).toFixed(2)} kg CO2` }
    ];
    
    const comparisonTbody = document.getElementById('comparison-data');
    comparisonTbody.innerHTML = '';
    comparisonData.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.metric}</td>
            <td>${item.before}</td>
            <td>${item.after}</td>
            <td>${item.savings}</td>
        `;
        comparisonTbody.appendChild(row);
    });
    
    // Display appliance details
    const applianceTbody = document.getElementById('appliance-details');
    applianceTbody.innerHTML = '';
    applianceDetails.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.name}</td>
            <td>${item.qty}</td>
            <td>${item.hours}</td>
            <td>${item.monthlyUsage}</td>
        `;
        applianceTbody.appendChild(row);
    });
    
    // Create/update chart
    createCostChart(totalBill, newBill);
    
    // Store current calculation data for potential saving
    window.currentCalculation = {
        houseNumber,
        meterNumber,
        location,
        weather,
        costPerKwh: costValue,
        totalConsumption,
        totalBill,
        renewableSource,
        installationCost,
        newBill,
        monthlySavings,
        yearlySavings,
        applianceDetails,
        timestamp: new Date().toLocaleString()
    };
    
    // Switch to result interface
    switchInterface('result');
}

function getRenewableRecommendation(location, weather, consumption) {
    let renewableSource = '';
    let installationCost = 0;
    let reductionPercentage = 0;
    let suggestions = [];
    let benefits = [];
    let considerations = [];

    // Determine recommendation based on weather and location
    if ((weather === 'sunny' || weather === 'mostly-sunny') && (location === 'rural' || location === 'coastal')) {
        renewableSource = 'Solar Panels (Highly Recommended)';
        installationCost = Math.round(consumption * 20000);
        reductionPercentage = 70;
        suggestions = renewableSuggestions.solar;
        benefits = [
            "Excellent sunlight availability in your area",
            "High return on investment (5-7 years)",
            "Low maintenance requirements",
            "30-40% government subsidies available",
            "Can sell excess power back to grid"
        ];
        considerations = [
            "Requires adequate roof space (10 sq.m per kW)",
            "Initial investment of ₹20,000-25,000 per kW",
            "May need occasional cleaning",
            "Output reduced on cloudy days",
            "Batteries needed for off-grid systems add cost"
        ];
    } 
    else if (weather === 'windy' && (location === 'hilly' || location === 'coastal')) {
        renewableSource = 'Wind Turbine (Recommended)';
        installationCost = Math.round(consumption * 30000);
        reductionPercentage = 60;
        suggestions = renewableSuggestions.wind;
        benefits = [
            "Consistent wind patterns in your area (avg. 5-8 m/s)",
            "Good performance year-round",
            "Can generate power day and night",
            "25-30% government subsidies available",
            "Can work alongside solar systems"
        ];
        considerations = [
            "Requires proper zoning permissions",
            "Needs minimum wind speed of 5 m/s",
            "Regular maintenance required",
            "Turbine noise may be consideration",
            "Higher initial cost than solar"
        ];
    }
    else if (weather === 'rainy' && location === 'hilly') {
        renewableSource = 'Micro-Hydro Power (Recommended)';
        installationCost = Math.round(consumption * 40000);
        reductionPercentage = 80;
        suggestions = renewableSuggestions.hydro;
        benefits = [
            "Very high efficiency in hilly areas",
            "Continuous 24/7 power generation",
            "Long system lifespan (30+ years)",
            "Minimal environmental impact",
            "Very low operating costs"
        ];
        considerations = [
            "Requires flowing water source (stream/river)",
            "Seasonal variations may affect output",
            "Initial civil work required",
            "Water rights may be needed",
            "Higher initial investment"
        ];
    }
    else if (location === 'urban' || weather === 'unpredictable') {
        renewableSource = 'Hybrid System (Solar + Wind)';
        installationCost = Math.round(consumption * 25000);
        reductionPercentage = 50;
        suggestions = renewableSuggestions.hybrid;
        benefits = [
            "Balanced energy production",
            "Good for variable weather conditions",
            "Reduces grid dependence",
            "Can combine with battery storage",
            "More consistent year-round output"
        ];
        considerations = [
            "Higher initial cost than single systems",
            "Requires more space for both systems",
            "More complex installation",
            "May need professional maintenance",
            "System sizing more critical"
        ];
    }
    else {
        // Default recommendation
        renewableSource = 'Solar Panels (Standard)';
        installationCost = Math.round(consumption * 18000);
        reductionPercentage = 40;
        suggestions = [
            "Standard solar panel installation",
            "Consider energy efficiency upgrades",
            "Explore government subsidies",
            "Start with smaller system and expand",
            "Grid-tied system recommended"
        ];
        benefits = [
            "Proven technology with good reliability",
            "Widely available components",
            "Easy to maintain and operate",
            "30% government subsidies available",
            "Scalable - can add more panels later"
        ];
        considerations = [
            "Depends on sunlight availability",
            "Requires proper orientation (south-facing)",
            "May need grid connection for net metering",
            "Output varies by season",
            "Roof condition must support panels"
        ];
    }

    // Add general suggestions
    suggestions = [...suggestions, ...renewableSuggestions.general];
    
    return { 
        renewableSource, 
        installationCost, 
        reductionPercentage, 
        suggestions,
        benefits,
        considerations
    };
}

function displayRenewableDetails(source, benefits, considerations) {
    const container = document.getElementById('renewable-details');
    container.innerHTML = `
        <div class="renewable-content">
            <div class="renewable-benefits">
                <h4>
                    <svg viewBox="0 0 24 24" width="18" height="18">
                        <path fill="#1b5e20" d="M12,2C6.48,2 2,6.48 2,12s4.48,10 10,10 10-4.48 10-10S17.52,2 12,2zm-2,15l-5-5 1.41-1.41L10,14.17l7.59-7.59L19,8l-9,9z"/>
                    </svg>
                    Benefits
                </h4>
                <ul>
                    ${benefits.map(b => `<li>${b}</li>`).join('')}
                </ul>
            </div>
            <div class="renewable-considerations">
                <h4>
                    <svg viewBox="0 0 24 24" width="18" height="18">
                        <path fill="#c62828" d="M11,15h2v2h-2v-2zm0-8h2v6h-2V7zm1-5C6.47,2 2,6.47 2,12s4.47,10 10,10 10-4.47 10-10S17.53,2 12,2zm0,18c-4.41,0-8-3.59-8-8s3.59-8 8-8 8,3.59 8,8-3.59,8-8,8z"/>
                    </svg>
                    Considerations
                </h4>
                <ul>
                    ${considerations.map(c => `<li>${c}</li>`).join('')}
                </ul>
            </div>
        </div>
    `;
}

function displayRotatingSuggestions(suggestions) {
    const display = document.getElementById('suggestion-display');
    display.innerHTML = `<p>${suggestions[0]}</p>`;
    
    let index = 1;
    clearInterval(suggestionInterval);
    suggestionInterval = setInterval(() => {
        display.innerHTML = `<p>${suggestions[index]}</p>`;
        index = (index + 1) % suggestions.length;
    }, 5000);
}

function createCostChart(beforeCost, afterCost) {
    const ctx = document.getElementById('costChart').getContext('2d');
    
    // Destroy previous chart if it exists
    if (costChart) {
        costChart.destroy();
    }
    
    costChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Monthly Cost', 'Yearly Cost'],
            datasets: [
                {
                    label: 'Before Optimization',
                    backgroundColor: '#ef5350',
                    data: [beforeCost, beforeCost * 12]
                },
                {
                    label: 'After Optimization',
                    backgroundColor: '#66bb6a',
                    data: [afterCost, afterCost * 12]
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Cost (₹)'
                    }
                }
            },
            plugins: {
                title: {
                    display: true,
                    text: 'Energy Cost Comparison',
                    font: {
                        size: 16
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return `${context.dataset.label}: ₹${context.raw.toFixed(2)}`;
                        }
                    }
                }
            }
        }
    });
}

function saveCurrentData() {
    playClickSound();
    
    if (!window.currentCalculation) {
        alert('No calculation data to save.');
        return;
    }
    
    // Add to app data with serial number
    const newEntry = {
        serial: appData.length + 1,
        ...window.currentCalculation
    };
    
    appData.push(newEntry);
    
    // Save to localStorage
    localStorage.setItem('energyOptimizationData', JSON.stringify(appData));
    
    alert('Data saved successfully! This will be available even when you return later.');
}

function checkPassword() {
    playClickSound();
    
    const password = document.getElementById('password').value;
    const correctPassword = '2327IILM';
    
    if (password === correctPassword) {
        switchInterface('dataView');
    } else {
        alert('Incorrect password. Please try again.');
    }
    
    // Clear password field
    document.getElementById('password').value = '';
}

function displaySavedData() {
    const tbody = document.getElementById('saved-data-body');
    tbody.innerHTML = '';
    
    appData.forEach(data => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${data.serial}</td>
            <td>${data.houseNumber}</td>
            <td>${data.meterNumber}</td>
            <td>${data.location}</td>
            <td>${data.weather}</td>
            <td>${data.totalConsumption.toFixed(2)}</td>
            <td>${data.totalBill.toFixed(2)}</td>
            <td>${data.renewableSource.split('(')[0].trim()}</td>
            <td>${data.installationCost.toLocaleString()}</td>
            <td>${data.newBill.toFixed(2)}</td>
            <td>${data.monthlySavings.toFixed(2)}</td>
            <td>${data.timestamp}</td>
        `;
        tbody.appendChild(row);
    });
}

function exportToExcel() {
    playClickSound();
    
    if (appData.length === 0) {
        alert('No data to export.');
        return;
    }
    
    // Prepare worksheet
    const wsData = [
        ['S.No', 'House No', 'Meter No', 'Location', 'Weather', 'Total Usage (kWh)', 
         'Bill Before (₹)', 'Renewable Source', 'Install Cost (₹)', 'New Bill (₹)', 
         'Savings (₹)', 'Timestamp']
    ];
    
    appData.forEach(data => {
        wsData.push([
            data.serial,
            data.houseNumber,
            data.meterNumber,
            data.location,
            data.weather,
            data.totalConsumption.toFixed(2),
            data.totalBill.toFixed(2),
            data.renewableSource.split('(')[0].trim(),
            data.installationCost,
            data.newBill.toFixed(2),
            data.monthlySavings.toFixed(2),
            data.timestamp
        ]);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'EnergyData');
    
    // Generate and download file
    const fileName = `Energy_Optimization_Data_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
}