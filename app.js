// Main Application Logic
// Handles Office.js integration and UI interactions

let spClient;
let currentEmail = {};
let currentPipeline = null;
let currentBoxes = [];
let allPipelines = [];
let allStages = {};

// Initialize Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Initialize SharePoint client
        // IMPORTANT: Replace with your actual SharePoint site URL
        spClient = new SharePointClient('https://yourtenant.sharepoint.com/sites/yoursite');
        
        // Load current email info
        loadEmailInfo();
        
        // Setup UI event listeners
        setupEventListeners();
        
        // Load initial data
        loadPipelines();
        loadRecentActivity();
        
        console.log('Drycake CRM initialized');
    }
});

// Load current email information
function loadEmailInfo() {
    const item = Office.context.mailbox.item;
    
    if (item) {
        // Get email details
        currentEmail = {
            subject: item.subject,
            messageId: item.internetMessageId || item.itemId,
            date: item.dateTimeCreated || new Date()
        };
        
        // Get sender information
        if (item.from) {
            currentEmail.from = item.from.emailAddress;
            currentEmail.fromName = item.from.displayName;
        } else {
            currentEmail.from = 'Unknown';
            currentEmail.fromName = 'Unknown';
        }
        
        // Get recipients
        if (item.to && item.to.length > 0) {
            currentEmail.to = item.to.map(r => r.emailAddress).join(', ');
        } else {
            currentEmail.to = '';
        }
        
        // Update UI
        document.getElementById('emailSubject').textContent = currentEmail.subject || 'No Subject';
        document.getElementById('emailFrom').textContent = `${currentEmail.fromName} <${currentEmail.from}>`;
        document.getElementById('emailDate').textContent = new Date(currentEmail.date).toLocaleString();
        
        // Check if this email is already linked to any boxes
        checkEmailLinks();
    }
}

// Check if current email is linked to boxes
async function checkEmailLinks() {
    try {
        const linkedBoxes = await spClient.getBoxesByEmailMessageId(currentEmail.messageId);
        
        if (linkedBoxes && linkedBoxes.length > 0) {
            displayLinkedBoxes(linkedBoxes);
        }
    } catch (error) {
        console.error('Error checking email links:', error);
    }
}

// Display linked boxes in the UI
function displayLinkedBoxes(boxes) {
    const container = document.getElementById('linkedBoxesList');
    
    if (!boxes || boxes.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">📦</div>
                <p>No boxes linked to this email</p>
            </div>
        `;
        return;
    }
    
    let html = '<ul class="box-list">';
    boxes.forEach(box => {
        html += `
            <li class="box-item" data-box-id="${box.ID}">
                <div class="box-name">${escapeHtml(box.Title)}</div>
                <div class="box-meta">
                    Pipeline: ${box.PipelineId} | Stage: ${box.StageId}
                    ${box.BoxValue > 0 ? ` | Value: $${box.BoxValue.toLocaleString()}` : ''}
                </div>
            </li>
        `;
    });
    html += '</ul>';
    
    container.innerHTML = html;
    
    // Add click handlers
    container.querySelectorAll('.box-item').forEach(item => {
        item.addEventListener('click', () => {
            const boxId = item.dataset.boxId;
            viewBoxDetails(boxId);
        });
    });
}

// Load all pipelines
async function loadPipelines() {
    try {
        allPipelines = await spClient.getPipelines();
        
        // Update all pipeline selectors
        updatePipelineSelectors();
        
        // Load boxes for the first pipeline
        if (allPipelines && allPipelines.length > 0) {
            currentPipeline = allPipelines[0].ID;
            document.getElementById('pipelineSelect').value = currentPipeline;
            loadBoxes(currentPipeline);
            loadStages(currentPipeline);
        }
    } catch (error) {
        console.error('Error loading pipelines:', error);
        showError('Failed to load pipelines. Please check your SharePoint connection.');
    }
}

// Update all pipeline select dropdowns
function updatePipelineSelectors() {
    const selectors = [
        document.getElementById('pipelineSelect'),
        document.getElementById('linkPipelineSelect'),
        document.getElementById('createPipelineSelect')
    ];
    
    selectors.forEach(select => {
        if (!select) return;
        
        select.innerHTML = allPipelines.map(p => 
            `<option value="${p.ID}">${escapeHtml(p.Title)}</option>`
        ).join('');
    });
}

// Load stages for a pipeline
async function loadStages(pipelineId) {
    try {
        const stages = await spClient.getStagesByPipeline(pipelineId);
        allStages[pipelineId] = stages;
        
        // Update stage selector in create modal
        const stageSelect = document.getElementById('stageSelect');
        if (stageSelect && stages && stages.length > 0) {
            stageSelect.innerHTML = stages.map(s => 
                `<option value="${s.ID}">${escapeHtml(s.Title)}</option>`
            ).join('');
        }
    } catch (error) {
        console.error('Error loading stages:', error);
    }
}

// Load boxes for a pipeline
async function loadBoxes(pipelineId) {
    try {
        const boxesList = document.getElementById('boxesList');
        boxesList.innerHTML = '<li class="loading">Loading boxes...</li>';
        
        currentBoxes = await spClient.getBoxesByPipeline(pipelineId);
        
        if (!currentBoxes || currentBoxes.length === 0) {
            boxesList.innerHTML = `
                <div class="empty-state">
                    <div class="empty-state-icon">📦</div>
                    <p>No boxes in this pipeline</p>
                </div>
            `;
            return;
        }
        
        // Get stages for display
        const stages = allStages[pipelineId] || [];
        const stageMap = {};
        stages.forEach(s => stageMap[s.ID] = s.Title);
        
        // Build boxes list
        let html = '';
        currentBoxes.forEach(box => {
            const stageName = stageMap[box.StageId] || 'Unknown Stage';
            html += `
                <li class="box-item" data-box-id="${box.ID}">
                    <div class="box-name">
                        ${escapeHtml(box.Title)}
                        <span class="stage-badge">${escapeHtml(stageName)}</span>
                    </div>
                    <div class="box-meta">
                        ${box.ContactName ? escapeHtml(box.ContactName) + ' | ' : ''}
                        ${box.BoxValue > 0 ? '$' + box.BoxValue.toLocaleString() : 'No value'}
                        | Updated: ${new Date(box.Modified).toLocaleDateString()}
                    </div>
                </li>
            `;
        });
        
        boxesList.innerHTML = html;
        
        // Add click handlers
        boxesList.querySelectorAll('.box-item').forEach(item => {
            item.addEventListener('click', () => {
                // Remove previous selection
                boxesList.querySelectorAll('.box-item').forEach(i => i.classList.remove('selected'));
                // Add selection to clicked item
                item.classList.add('selected');
            });
        });
        
        // Update link box select
        updateLinkBoxSelect(currentBoxes);
        
    } catch (error) {
        console.error('Error loading boxes:', error);
        showError('Failed to load boxes.');
    }
}

// Update the box selector in link modal
function updateLinkBoxSelect(boxes) {
    const select = document.getElementById('linkBoxSelect');
    if (!select || !boxes) return;
    
    select.innerHTML = '<option value="">Select box...</option>' + 
        boxes.map(b => `<option value="${b.ID}">${escapeHtml(b.Title)}</option>`).join('');
}

// Load recent activity
async function loadRecentActivity() {
    try {
        const activities = await spClient.getRecentActivities(20);
        const container = document.getElementById('activityList');
        
        if (!activities || activities.length === 0) {
            container.innerHTML = `
                <div class="empty-state">
                    <div class="empty-state-icon">📊</div>
                    <p>No recent activity</p>
                </div>
            `;
            return;
        }
        
        let html = '';
        activities.forEach(activity => {
            const author = activity.Author ? activity.Author.Title : 'Unknown';
            html += `
                <div class="activity-item">
                    <div class="activity-date">${new Date(activity.Created).toLocaleString()}</div>
                    <div class="activity-text">
                        <strong>${escapeHtml(author)}</strong> ${escapeHtml(activity.ActivityType)}: 
                        ${escapeHtml(activity.ActivityText)}
                    </div>
                </div>
            `;
        });
        
        container.innerHTML = html;
    } catch (error) {
        console.error('Error loading activity:', error);
    }
}

// Setup event listeners
function setupEventListeners() {
    // Tab navigation
    document.querySelectorAll('.tab-button').forEach(btn => {
        btn.addEventListener('click', (e) => {
            // Update active tab button
            document.querySelectorAll('.tab-button').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
            
            // Show corresponding tab content
            const tabName = e.target.dataset.tab;
            document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
            document.getElementById(`${tabName}-tab`).classList.add('active');
        });
    });
    
    // Pipeline selector
    document.getElementById('pipelineSelect').addEventListener('change', (e) => {
        currentPipeline = parseInt(e.target.value);
        loadBoxes(currentPipeline);
        loadStages(currentPipeline);
    });
    
    // Link box button
    document.getElementById('linkBoxBtn').addEventListener('click', openLinkBoxModal);
    
    // Create box button
    document.getElementById('createBoxBtn').addEventListener('click', openCreateBoxModal);
    
    // New box button
    document.getElementById('newBoxBtn').addEventListener('click', openCreateBoxModal);
    
    // Refresh button
    document.getElementById('refreshBtn').addEventListener('click', () => {
        loadPipelines();
        loadRecentActivity();
        checkEmailLinks();
    });
    
    // Link box modal handlers
    document.getElementById('cancelLinkBtn').addEventListener('click', closeLinkBoxModal);
    document.getElementById('confirmLinkBtn').addEventListener('click', linkEmailToBox);
    document.getElementById('linkPipelineSelect').addEventListener('change', (e) => {
        const pipelineId = parseInt(e.target.value);
        if (pipelineId) {
            spClient.getBoxesByPipeline(pipelineId).then(boxes => {
                updateLinkBoxSelect(boxes);
            });
        }
    });
    
    // Create box modal handlers
    document.getElementById('cancelCreateBtn').addEventListener('click', closeCreateBoxModal);
    document.getElementById('confirmCreateBtn').addEventListener('click', createNewBox);
    document.getElementById('createPipelineSelect').addEventListener('change', (e) => {
        const pipelineId = parseInt(e.target.value);
        if (pipelineId) {
            loadStages(pipelineId);
        }
    });
}

// Modal functions
function openLinkBoxModal() {
    document.getElementById('linkBoxModal').classList.add('active');
    if (currentPipeline && currentBoxes) {
        document.getElementById('linkPipelineSelect').value = currentPipeline;
        updateLinkBoxSelect(currentBoxes);
    }
}

function closeLinkBoxModal() {
    document.getElementById('linkBoxModal').classList.remove('active');
}

function openCreateBoxModal() {
    document.getElementById('createBoxModal').classList.add('active');
    
    // Pre-fill with email info
    const subject = currentEmail.subject || '';
    document.getElementById('boxNameInput').value = subject;
    
    if (currentPipeline) {
        document.getElementById('createPipelineSelect').value = currentPipeline;
        loadStages(currentPipeline);
    }
}

function closeCreateBoxModal() {
    document.getElementById('createBoxModal').classList.remove('active');
    // Clear form
    document.getElementById('boxNameInput').value = '';
    document.getElementById('boxValueInput').value = '';
    document.getElementById('boxNotesInput').value = '';
}

// Link email to existing box
async function linkEmailToBox() {
    try {
        const boxId = parseInt(document.getElementById('linkBoxSelect').value);
        
        if (!boxId) {
            showError('Please select a box');
            return;
        }
        
        await spClient.linkEmailToBox(currentEmail, boxId);
        
        // Create activity
        await spClient.createActivity(
            boxId,
            'Email',
            `Email linked: ${currentEmail.subject}`
        );
        
        showSuccess('Email linked to box successfully!');
        closeLinkBoxModal();
        checkEmailLinks();
        loadRecentActivity();
        
    } catch (error) {
        console.error('Error linking email:', error);
        showError('Failed to link email to box');
    }
}

// Create new box
async function createNewBox() {
    try {
        const pipelineId = parseInt(document.getElementById('createPipelineSelect').value);
        const stageId = parseInt(document.getElementById('stageSelect').value);
        const boxName = document.getElementById('boxNameInput').value.trim();
        const boxValue = parseFloat(document.getElementById('boxValueInput').value) || 0;
        const notes = document.getElementById('boxNotesInput').value.trim();
        
        if (!pipelineId || !stageId || !boxName) {
            showError('Please fill in all required fields');
            return;
        }
        
        const boxData = {
            name: boxName,
            pipelineId: pipelineId,
            stageId: stageId,
            value: boxValue,
            contactEmail: currentEmail.from,
            contactName: currentEmail.fromName,
            notes: notes
        };
        
        const newBox = await spClient.createBox(boxData);
        
        // Link current email to the new box
        await spClient.linkEmailToBox(currentEmail, newBox.ID);
        
        // Create activity
        await spClient.createActivity(
            newBox.ID,
            'Created',
            `Box created from email: ${currentEmail.subject}`
        );
        
        showSuccess('Box created successfully!');
        closeCreateBoxModal();
        loadBoxes(pipelineId);
        checkEmailLinks();
        loadRecentActivity();
        
    } catch (error) {
        console.error('Error creating box:', error);
        showError('Failed to create box');
    }
}

// View box details (future enhancement)
function viewBoxDetails(boxId) {
    // TODO: Implement detailed box view
    console.log('View box details:', boxId);
}

// Utility functions
function showError(message) {
    // Simple error notification - you can enhance this
    const content = document.querySelector('.content');
    const errorDiv = document.createElement('div');
    errorDiv.className = 'error';
    errorDiv.textContent = message;
    content.insertBefore(errorDiv, content.firstChild);
    
    setTimeout(() => errorDiv.remove(), 5000);
}

function showSuccess(message) {
    // Simple success notification - you can enhance this
    const content = document.querySelector('.content');
    const successDiv = document.createElement('div');
    successDiv.className = 'success';
    successDiv.textContent = message;
    content.insertBefore(successDiv, content.firstChild);
    
    setTimeout(() => successDiv.remove(), 3000);
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
