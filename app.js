// MetricFlow Hackathon - Advanced App Logic
document.addEventListener('DOMContentLoaded', () => {
    
    // --- State Management ---
    let activeDbKey = 'metricflow_db_v2';
    
    let db = { goals: [], audit: [], webhooks: [], activeCycle: 'phase1' };
    
    const loadDb = (key) => {
        activeDbKey = key;
        db = JSON.parse(localStorage.getItem(activeDbKey)) || {
            goals: [], audit: [], webhooks: [], activeCycle: 'phase1'
        };
    };

    const saveDb = () => localStorage.setItem(activeDbKey, JSON.stringify(db));

    const logAudit = (action, details, user) => {
        db.audit.unshift({ id: Date.now().toString(), timestamp: new Date().toISOString(), action, details, user });
        saveDb();
    };

    // --- DOM Elements ---
    const loginRoleSelect = document.getElementById('loginRoleSelect');
    const navUserDisplay = document.getElementById('navUserDisplay');
    const viewContainer = document.getElementById('viewContainer');
    const toastContainer = document.getElementById('toast-container');
    const loginScreen = document.getElementById('loginScreen');
    const appWrapper = document.getElementById('app');
    const integrationHub = document.getElementById('integrationHub');
    const hubContent = document.getElementById('hubContent');
    const notifBadge = document.getElementById('notifBadge');
    
    let currentUser = null;

    // --- Theme Logic ---
    const btnThemeToggle = document.getElementById('btnThemeToggle');
    let isDarkMode = false;
    btnThemeToggle?.addEventListener('click', () => {
        isDarkMode = !isDarkMode;
        document.body.classList.toggle('dark-mode', isDarkMode);
        btnThemeToggle.innerText = isDarkMode ? '☀️' : '🌙';
        // Re-render chart colors if active
        if (typeof Chart !== 'undefined') {
            Chart.defaults.color = isDarkMode ? '#94a3b8' : '#475569';
            Chart.defaults.borderColor = isDarkMode ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)';
        }
    });

    // --- Helpers ---
    const showToast = (message, type = 'info') => {
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.innerText = message;
        toastContainer.appendChild(toast);
        setTimeout(() => {
            toast.style.animation = 'fadeIn 0.3s ease reverse forwards';
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    };

    const generateId = () => Math.random().toString(36).substr(2, 9);
    
    const calculateProgress = (goal) => {
        if (goal.achStatus === 'Completed') return (100).toFixed(1);
        if (goal.uom === 'Timeline') return (goal.achStatus === 'On Track' ? 50 : 0).toFixed(1);

        if (goal.actual === undefined || goal.actual === null || goal.actual === '') return (0).toFixed(1);
        
        let progress = 0;
        const target = parseFloat(goal.target);
        const actual = parseFloat(goal.actual);

        switch(goal.uom) {
            case 'Min': 
                progress = target === 0 ? (actual >= 0 ? 100 : 0) : (actual / target) * 100; 
                break;
            case 'Max': 
                if (target === 0 && actual === 0) progress = 100;
                else if (actual === 0) progress = 100;
                else progress = (target / actual) * 100; 
                break;
            case 'Zero': 
                progress = actual === 0 ? 100 : 0; 
                break;
        }
        return Math.min(Math.max(progress, 0), 100).toFixed(1);
    };

    const getCycleName = (code) => {
        const map = { phase1: 'Phase 1: Goal Setting', q1: 'Q1 Check-in', q2: 'Q2 Check-in', q3: 'Q3 Check-in', q4: 'Q4 / Annual' };
        return map[code] || code;
    };

    let currentUserRole = 'employee'; // Default

    // --- MSAL Configuration ---
    const msalConfig = {
        auth: {
            clientId: "465fc293-1cff-4ccf-b9b4-5206d6a2a974", // Your Client ID from the screenshot
            authority: "https://login.microsoftonline.com/common", // Changed from specific tenant ID to 'common' for personal accounts
            redirectUri: "http://localhost:5500/", // Let's try localhost instead of 127.0.0.1
        },
        cache: {
            cacheLocation: "sessionStorage",
            storeAuthStateInCookie: false,
        }
    };
    
    let msalInstance;
    try {
        msalInstance = new msal.PublicClientApplication(msalConfig);
    } catch (e) {
        console.error("MSAL Initialization Error:", e);
    }

    // --- MS Entra ID Login ---
    document.getElementById('btnEntraLogin').addEventListener('click', async () => {
        document.getElementById('howItWorksModal')?.classList.add('hidden');
        currentUserRole = loginRoleSelect.value;
        const roleLabels = { employee: 'Employee', manager: 'Manager', admin: 'Admin / HR' };

        try {
            const loginRequest = {
                scopes: ["User.Read"]
            };
            
            // Trigger real Microsoft Entra ID Login Popup
            const loginResponse = await msalInstance.loginPopup(loginRequest);
            const account = loginResponse.account;
            
            // Store the real user info
            currentUser = account;
            
            // Display real name alongside simulated role with Avatar
            const avatarUrl = `https://ui-avatars.com/api/?name=${encodeURIComponent(account.name)}&background=3b82f6&color=fff&rounded=true&size=32`;
            navUserDisplay.innerHTML = `<img src="${avatarUrl}" alt="Avatar" class="profile-avatar"> <span class="profile-name">${account.name} (${roleLabels[currentUserRole]})</span>`;

            loadDb('metricflow_db_v2');

            loginScreen.style.opacity = '0';
            setTimeout(() => {
                loginScreen.style.display = 'none';
                appWrapper.classList.remove('hidden');
                renderView();
                
                showToast(`Welcome back, ${account.name}! Auth successful.`, 'success');
                logAudit('System Login', `User ${account.username} authenticated via Entra ID`, roleLabels[currentUserRole]);
            }, 300);
            
        } catch (error) {
            console.error("Entra ID Login failed:", error);
            showToast("Microsoft Authentication failed or was cancelled.", "error");
        }
    });

    // --- Demo Fallback Login ---
    document.getElementById('btnDemoLogin').addEventListener('click', () => {
        document.getElementById('howItWorksModal')?.classList.add('hidden');
        currentUserRole = loginRoleSelect.value;
        const roleLabels = { employee: 'Employee', manager: 'Manager', admin: 'Admin / HR' };
        
        currentUser = {
            name: "Demo " + roleLabels[currentUserRole],
            username: "demo." + currentUserRole + "@metricflow.local"
        };
        
        loadDb('metricflow_demo_db');
        
        // Display name with Avatar
        const avatarUrl = `https://ui-avatars.com/api/?name=${encodeURIComponent(currentUser.name)}&background=8b5cf6&color=fff&rounded=true&size=32`;
        navUserDisplay.innerHTML = `<img src="${avatarUrl}" alt="Avatar" class="profile-avatar"> <span class="profile-name">${currentUser.name}</span>`;
        
        loginScreen.style.opacity = '0';
        setTimeout(() => {
            loginScreen.style.display = 'none';
            appWrapper.classList.remove('hidden');
            renderView();
            showToast('Logged in via Demo Mode');
        }, 500);
    });

    document.getElementById('btnLogout').addEventListener('click', () => {
        appWrapper.classList.add('hidden');
        loginScreen.style.display = 'flex';
        loginScreen.style.opacity = '1';
    });

    // --- How It Works Modal ---
    const howItWorksModal = document.getElementById('howItWorksModal');
    document.getElementById('linkHowItWorks').addEventListener('click', (e) => {
        e.preventDefault();
        howItWorksModal.classList.remove('hidden');
    });
    document.getElementById('btnCloseHowItWorks').addEventListener('click', () => howItWorksModal.classList.add('hidden'));
    document.getElementById('btnGotItHowItWorks').addEventListener('click', () => howItWorksModal.classList.add('hidden'));

    // --- Webhook / Integration Hub ---
    let unreadNotifs = 0;
    const triggerWebhook = (source, message) => {
        const id = generateId();
        const time = new Date().toLocaleTimeString();
        db.webhooks.unshift({ id, source, message, time });
        saveDb();
        
        showToast(`${source}: ${message}`, 'info');
        unreadNotifs++;
        notifBadge.innerText = unreadNotifs;
        notifBadge.classList.remove('hidden');
        renderWebhooks();
    };

    const renderWebhooks = () => {
        if (db.webhooks.length === 0) {
            hubContent.innerHTML = '<p class="text-muted" style="text-align:center; margin-top:2rem;">No new notifications</p>';
            return;
        }
        hubContent.innerHTML = db.webhooks.map(w => `
            <div class="notif-card">
                <div style="font-size:0.75rem; color:var(--text-muted); margin-bottom:0.25rem; display:flex; justify-content:space-between;">
                    <strong>${w.source}</strong> <span>${w.time}</span>
                </div>
                <div style="font-size:0.85rem;">${w.message}</div>
            </div>
        `).join('');
    };

    document.getElementById('btnShowIntegrations').addEventListener('click', () => {
        integrationHub.classList.toggle('transform-out');
        unreadNotifs = 0;
        notifBadge.classList.add('hidden');
    });
    document.getElementById('btnCloseHub').addEventListener('click', () => integrationHub.classList.add('transform-out'));

    // --- View Rendering ---
    const renderView = () => {
        const role = currentUserRole;
        document.getElementById('activeCycleBadge').innerText = getCycleName(db.activeCycle);
        
        viewContainer.innerHTML = '';
        viewContainer.classList.remove('fade-in');
        void viewContainer.offsetWidth; 
        viewContainer.classList.add('fade-in');
        setTimeout(() => viewContainer.classList.remove('fade-in'), 450);

        const template = document.getElementById(`tpl-${role}`).content.cloneNode(true);
        viewContainer.appendChild(template);

        if (role === 'employee') initEmployeeView();
        else if (role === 'manager') initManagerView();
        else if (role === 'admin') initAdminView();
    };

    // ==========================================
    // EMPLOYEE VIEW
    // ==========================================
    const initEmployeeView = () => {
        const employeeId = 'emp_1'; 
        let currentEditingGoalId = null;
        
        const isPhase1 = db.activeCycle === 'phase1';
        document.getElementById('empCycleText').innerText = `Current Window: ${getCycleName(db.activeCycle)}`;
        
        const warning = document.getElementById('cycleWarningEmp');
        if (!isPhase1) {
            warning.innerHTML = '<strong>Goal Setting is Locked:</strong> The active window is ' + getCycleName(db.activeCycle) + '. You can only update actual achievements.';
            warning.classList.remove('hidden');
            document.getElementById('btnNewGoal').disabled = true;
            document.getElementById('btnSubmitGoals').disabled = true;
        }

        const renderGoals = () => {
            const list = document.getElementById('employeeGoalList');
            const totalDisplay = document.getElementById('totalWeightDisplay');
            const btnSubmit = document.getElementById('btnSubmitGoals');
            list.innerHTML = '';

            const myGoals = db.goals.filter(g => g.ownerId === employeeId);
            let totalWeight = 0;
            myGoals.forEach(g => totalWeight += parseFloat(g.weightage));

            totalDisplay.innerText = `${totalWeight}%`;
            totalDisplay.className = `stat-value ${totalWeight === 100 ? 'success' : (totalWeight > 100 ? 'error' : '')}`;
            
            const isDraft = myGoals.some(g => g.status === 'Draft' || g.status === 'Returned');
            if (isPhase1) {
                btnSubmit.disabled = !(isDraft && Math.abs(totalWeight - 100) < 0.01 && myGoals.length > 0 && myGoals.length <= 8);
            }

            if (myGoals.length === 0) {
                list.innerHTML = `<div class="glass-panel" style="padding: 2rem; text-align: center; grid-column: 1/-1;"><p class="text-muted">No goals found.</p></div>`;
                return;
            }

            myGoals.forEach(goal => {
                const isLocked = ['Submitted', 'Approved'].includes(goal.status);
                const progress = calculateProgress(goal);
                
                const card = document.createElement('div');
                card.className = `goal-card glass-panel status-${goal.status.toLowerCase()} ${goal.isShared ? 'is-shared' : ''}`;
                
                let actionsHtml = '';
                if (!isLocked && isPhase1) {
                    actionsHtml = `
                        <button class="btn-secondary btn-sm btn-edit-goal" data-id="${goal.id}">Edit</button>
                        ${!goal.isShared ? `<button class="btn-danger btn-sm btn-delete-goal" data-id="${goal.id}">Delete</button>` : ''}
                    `;
                } else if (goal.status === 'Approved' && !isPhase1) {
                    actionsHtml = `<button class="btn-primary btn-sm btn-update-ach" data-id="${goal.id}">Update Progress</button>`;
                }

                card.innerHTML = `
                    <div class="goal-header">
                        <div>
                            <span class="goal-thrust">${goal.thrustArea}</span>
                            <h4 class="goal-title">${goal.title}</h4>
                        </div>
                        <span class="badge ${goal.status.toLowerCase()}">${goal.status} ${goal.isShared ? '(Shared)' : ''}</span>
                    </div>
                    ${goal.desc ? `<p class="text-muted" style="font-size: 0.85rem; margin-bottom: 1rem;">${goal.desc}</p>` : ''}
                    <div class="goal-meta">
                        <div class="meta-item"><span class="meta-label">UoM</span><span>${goal.uom}</span></div>
                        <div class="meta-item"><span class="meta-label">Weightage</span><span>${goal.weightage}%</span></div>
                        <div class="meta-item"><span class="meta-label">Target</span><span>${goal.uom === 'Timeline' ? goal.deadline : goal.target}</span></div>
                        <div class="meta-item"><span class="meta-label">Actual</span><span>${(goal.actual !== undefined && goal.actual !== null && goal.actual !== '') ? goal.actual : '-'}</span></div>
                    </div>
                    ${goal.status === 'Approved' ? `
                        <div class="progress-container">
                            <div class="progress-header"><span>Progress Score</span><span>${progress}%</span></div>
                            <div class="progress-track"><div class="progress-fill ${progress >= 100 ? 'bg-success' : ''}" style="width: ${progress}%"></div></div>
                        </div>
                    ` : ''}
                    ${goal.feedback ? `<div style="margin-top:0.5rem; font-size:0.8rem; color:var(--warning); background:rgba(245,158,11,0.1); padding:0.5rem; border-radius:4px;"><strong>Manager Note:</strong> ${goal.feedback}</div>` : ''}
                    <div class="goal-actions">${actionsHtml}</div>
                `;
                list.appendChild(card);
            });

            document.querySelectorAll('.btn-edit-goal').forEach(btn => btn.addEventListener('click', (e) => openGoalModal(e.target.dataset.id)));
            document.querySelectorAll('.btn-delete-goal').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    db.goals = db.goals.filter(g => g.id !== e.target.dataset.id); saveDb(); renderGoals(); showToast('Goal deleted');
                });
            });
            document.querySelectorAll('.btn-update-ach').forEach(btn => btn.addEventListener('click', (e) => openAchModal(e.target.dataset.id)));
        };

        const modal = document.getElementById('goalModal');
        const form = document.getElementById('goalForm');
        
        const openGoalModal = (goalId = null) => {
            currentEditingGoalId = goalId; form.reset();
            document.getElementById('targetWrapper').style.display = 'block'; document.getElementById('deadlineWrapper').style.display = 'none';
            document.getElementById('goalTarget').setAttribute('required', 'true'); document.getElementById('goalDeadline').removeAttribute('required');

            if (goalId) {
                const goal = db.goals.find(g => g.id === goalId);
                document.getElementById('goalThrustArea').value = goal.thrustArea;
                document.getElementById('goalTitle').value = goal.title;
                document.getElementById('goalDesc').value = goal.desc || '';
                document.getElementById('goalUom').value = goal.uom;
                document.getElementById('goalWeightage').value = goal.weightage;
                
                if (goal.uom === 'Timeline') {
                    document.getElementById('targetWrapper').style.display = 'none';
                    document.getElementById('deadlineWrapper').style.display = 'block';
                    document.getElementById('goalDeadline').value = goal.deadline || '';
                    document.getElementById('goalTarget').removeAttribute('required'); document.getElementById('goalDeadline').setAttribute('required', 'true');
                } else document.getElementById('goalTarget').value = goal.target;

                const isShared = goal.isShared;
                document.getElementById('goalTitle').readOnly = isShared;
                document.getElementById('goalTarget').readOnly = isShared;
                document.getElementById('goalThrustArea').disabled = isShared;
                document.getElementById('goalUom').disabled = isShared;
            } else {
                document.getElementById('goalTitle').readOnly = false; document.getElementById('goalTarget').readOnly = false;
                document.getElementById('goalThrustArea').disabled = false; document.getElementById('goalUom').disabled = false;
            }
            modal.classList.remove('hidden');
        };

        const closeGoalModal = () => modal.classList.add('hidden');
        document.getElementById('btnNewGoal').addEventListener('click', () => { if (db.goals.filter(g => g.ownerId === employeeId).length >= 8) showToast('Max 8 goals allowed.', 'error'); else openGoalModal(); });
        document.getElementById('btnCloseGoalModal').addEventListener('click', closeGoalModal); document.getElementById('btnCancelGoal').addEventListener('click', closeGoalModal);
        document.getElementById('goalUom').addEventListener('change', (e) => {
            if (e.target.value === 'Timeline') {
                document.getElementById('targetWrapper').style.display = 'none'; document.getElementById('deadlineWrapper').style.display = 'block';
                document.getElementById('goalTarget').removeAttribute('required'); document.getElementById('goalDeadline').setAttribute('required', 'true');
            } else {
                document.getElementById('targetWrapper').style.display = 'block'; document.getElementById('deadlineWrapper').style.display = 'none';
                document.getElementById('goalTarget').setAttribute('required', 'true'); document.getElementById('goalDeadline').removeAttribute('required');
            }
        });

        // AI Optimizer
        const btnAiImprove = document.getElementById('btnAiImprove');
        btnAiImprove?.addEventListener('click', () => {
            const titleInput = document.getElementById('goalTitle');
            const descInput = document.getElementById('goalDesc');
            if (!titleInput.value) return showToast("Please enter a basic goal title first so AI has context!", "warning");
            
            btnAiImprove.classList.add('thinking');
            btnAiImprove.innerText = '✨ Thinking...';
            
            setTimeout(() => {
                const aiSuggestions = [
                    { title: "Optimize operational pipeline by 15% to reduce overhead", desc: "Evaluate current procedural bottlenecks and implement automated workflows to achieve a 15% reduction in total processing time by end of Q3." },
                    { title: "Increase net-new customer acquisition by 20% through targeted outreach", desc: "Leverage data-driven marketing campaigns and A/B testing on primary landing pages to secure a 20% bump in qualified leads." },
                    { title: "Enhance system reliability to 99.99% uptime", desc: "Refactor core legacy modules and implement redundant load-balancing to ensure maximum availability during peak traffic hours." }
                ];
                const pick = aiSuggestions[Math.floor(Math.random() * aiSuggestions.length)];
                titleInput.value = pick.title; descInput.value = pick.desc;
                
                btnAiImprove.classList.remove('thinking');
                btnAiImprove.innerText = '✨ AI Optimize';
                showToast("Goal successfully optimized using S.M.A.R.T framework!", "success");
            }, 1500);
        });

        document.getElementById('btnSaveGoal').addEventListener('click', () => {
            if (!form.checkValidity()) { form.reportValidity(); return; }
            const weightage = parseFloat(document.getElementById('goalWeightage').value);
            if (weightage < 10) return showToast('Min 10% weightage per goal.', 'error');

            const myGoals = db.goals.filter(g => g.ownerId === employeeId);
            let currentTotal = 0;
            myGoals.forEach(g => {
                if (g.id !== currentEditingGoalId) currentTotal += parseFloat(g.weightage);
            });
            
            if (currentTotal + weightage > 100) {
                return showToast(`Total weightage cannot exceed 100%. You can add up to ${100 - currentTotal}% more.`, 'error');
            }
            
            if (!currentEditingGoalId && myGoals.length >= 8) {
                return showToast('Maximum 8 goals allowed.', 'error');
            }

            const goalData = {
                thrustArea: document.getElementById('goalThrustArea').value, title: document.getElementById('goalTitle').value, desc: document.getElementById('goalDesc').value, uom: document.getElementById('goalUom').value, weightage: weightage,
                target: document.getElementById('goalUom').value === 'Timeline' ? '-' : document.getElementById('goalTarget').value, deadline: document.getElementById('goalUom').value === 'Timeline' ? document.getElementById('goalDeadline').value : null,
            };

            if (currentEditingGoalId) {
                const idx = db.goals.findIndex(g => g.id === currentEditingGoalId);
                db.goals[idx] = { ...db.goals[idx], ...goalData, status: 'Draft' };
            } else {
                db.goals.push({ id: generateId(), ownerId: employeeId, status: 'Draft', isShared: false, ...goalData });
            }

            saveDb(); closeGoalModal(); renderGoals(); showToast('Goal saved.', 'success');
        });

        document.getElementById('btnSubmitGoals').addEventListener('click', () => {
            const myGoals = db.goals.filter(g => g.ownerId === employeeId);
            myGoals.forEach(g => { if (g.status === 'Draft' || g.status === 'Returned') g.status = 'Submitted'; });
            saveDb(); logAudit('Goals Submitted', `Employee ${employeeId} submitted goals`, 'Employee');
            triggerWebhook('MS Teams', `Notification sent to Manager: John Doe has submitted their Phase 1 Goal Sheet.`);
            renderGoals(); showToast('Goals submitted to Manager.', 'success');
        });

        const achModal = document.getElementById('achievementModal');
        let currentAchGoalId = null;
        const openAchModal = (goalId) => {
            currentAchGoalId = goalId; const goal = db.goals.find(g => g.id === goalId);
            document.getElementById('achStatus').value = goal.achStatus || 'Not Started'; 
            document.getElementById('achActual').value = (goal.actual !== undefined && goal.actual !== null) ? goal.actual : '';
            document.getElementById('actualWrapper').style.display = goal.uom === 'Timeline' ? 'none' : 'block';
            achModal.classList.remove('hidden');
        };

        document.getElementById('btnCloseAchModal').addEventListener('click', () => achModal.classList.add('hidden')); document.getElementById('btnCancelAch').addEventListener('click', () => achModal.classList.add('hidden'));
        document.getElementById('btnSaveAch').addEventListener('click', () => {
            const formAch = document.getElementById('achForm'); if (!formAch.checkValidity()) return formAch.reportValidity();
            const goal = db.goals.find(g => g.id === currentAchGoalId);
            
            const prevProgress = calculateProgress(goal);
            
            goal.achStatus = document.getElementById('achStatus').value;
            if (goal.uom !== 'Timeline') goal.actual = document.getElementById('achActual').value;
            saveDb(); achModal.classList.add('hidden'); renderGoals(); 
            
            const newProgress = calculateProgress(goal);
            if (newProgress == 100 && prevProgress != 100) {
                if (typeof confetti === 'function') confetti({ particleCount: 150, spread: 70, origin: { y: 0.6 } });
                showToast('Goal 100% Completed! Awesome job!', 'success');
            } else {
                showToast('Progress updated.', 'success');
            }
        });

        renderGoals();
    };

    // ==========================================
    // MANAGER VIEW
    // ==========================================
    const initManagerView = () => {
        const mgrContent = document.getElementById('mgrContentArea');
        const feedbackModal = document.getElementById('mgrFeedbackModal');
        let actionGoalId = null; let actionType = null;
        
        const isPhase1 = db.activeCycle === 'phase1';
        if (!isPhase1) {
            const w = document.getElementById('cycleWarningMgr');
            w.innerHTML = `<strong>Approvals Locked:</strong> The active window is ${getCycleName(db.activeCycle)}. Only Quarterly Check-ins are enabled.`;
            w.classList.remove('hidden');
        }

        const renderApprovals = () => {
            const submittedGoals = db.goals.filter(g => g.status === 'Submitted');
            if (submittedGoals.length === 0) return mgrContent.innerHTML = `<div class="glass-panel" style="padding: 2rem; text-align: center;"><p class="text-muted">No goals pending approval.</p></div>`;

            let html = `<div class="table-container"><table><thead><tr><th>Employee</th><th>Title</th><th>Weight</th><th>Target</th><th>Actions</th></tr></thead><tbody>`;
            submittedGoals.forEach(g => {
                html += `<tr><td>John Doe</td><td><strong>${g.title}</strong><br><span style="font-size: 0.75rem; color: var(--text-muted);">${g.thrustArea}</span></td>
                    <td><div class="inline-input-group"><input type="number" value="${g.weightage}" id="weight_${g.id}" class="inline-edit" ${!isPhase1?'disabled':''}><span>%</span></div></td>
                    <td>${g.uom === 'Timeline' ? `<input type="date" value="${g.deadline}" id="target_${g.id}" class="inline-edit" ${!isPhase1?'disabled':''}>` : `<input type="number" value="${g.target}" id="target_${g.id}" class="inline-edit" ${!isPhase1?'disabled':''}>`}</td>
                    <td>
                        <button class="btn-success btn-sm btn-approve" data-id="${g.id}" ${!isPhase1?'disabled':''}>Approve</button>
                        <button class="btn-danger btn-sm btn-return" data-id="${g.id}" ${!isPhase1?'disabled':''}>Return</button>
                    </td></tr>`;
            });
            html += `</tbody></table></div>`; mgrContent.innerHTML = html;

            document.querySelectorAll('.btn-approve').forEach(btn => btn.addEventListener('click', (e) => { saveInlineEdits(e.target.dataset.id); openFeedbackModal(e.target.dataset.id, 'approve'); }));
            document.querySelectorAll('.btn-return').forEach(btn => btn.addEventListener('click', (e) => { saveInlineEdits(e.target.dataset.id); openFeedbackModal(e.target.dataset.id, 'return'); }));
        };

        const saveInlineEdits = (id) => {
            const goal = db.goals.find(g => g.id === id); const wInput = document.getElementById(`weight_${id}`); const tInput = document.getElementById(`target_${id}`);
            if (wInput) goal.weightage = parseFloat(wInput.value);
            if (tInput) { if(goal.uom === 'Timeline') goal.deadline = tInput.value; else goal.target = tInput.value; }
        };

        const openFeedbackModal = (id, type) => {
            actionGoalId = id; actionType = type;
            document.getElementById('mgrFeedbackTitle').innerText = type === 'approve' ? 'Approve Goal' : 'Return for Rework';
            document.getElementById('mgrFeedbackText').value = ''; feedbackModal.classList.remove('hidden');
        };

        document.getElementById('btnCloseFeedbackModal').addEventListener('click', () => feedbackModal.classList.add('hidden')); document.getElementById('btnCancelFeedback').addEventListener('click', () => feedbackModal.classList.add('hidden'));
        document.getElementById('btnSubmitFeedback').addEventListener('click', () => {
            const feedback = document.getElementById('mgrFeedbackText').value;
            if(!feedback.trim()) return showToast('Please provide feedback.', 'error');
            const goal = db.goals.find(g => g.id === actionGoalId);
            goal.feedback = feedback;
            if (actionType === 'approve') {
                goal.status = 'Approved'; logAudit('Goal Approved', `Goal ${goal.id} approved`, 'Manager');
                triggerWebhook('Outlook Email', `To: John Doe. Subject: Goal Approved. Your manager has approved your goal: ${goal.title}`);
            } else {
                goal.status = 'Returned'; logAudit('Goal Returned', `Goal ${goal.id} returned`, 'Manager');
                triggerWebhook('MS Teams', `Alert to John Doe: Your goal requires rework. Note: ${feedback}`);
            }
            saveDb(); feedbackModal.classList.add('hidden'); renderApprovals();
        });

        // Tabs
        document.querySelectorAll('.tab-btn').forEach(btn => btn.addEventListener('click', (e) => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active')); e.target.classList.add('active');
            e.target.dataset.tab === 'approvals' ? renderApprovals() : renderCheckins();
        }));

        const renderCheckins = () => {
            const approvedGoals = db.goals.filter(g => g.status === 'Approved');
            if (approvedGoals.length === 0) return mgrContent.innerHTML = `<div class="glass-panel" style="padding: 2rem; text-align: center;"><p class="text-muted">No approved goals available.</p></div>`;

            let html = `<div class="goal-grid">`;
            approvedGoals.forEach((g, idx) => {
                const progress = calculateProgress(g);
                html += `
                    <div class="goal-card glass-panel stagger-${(idx % 4) + 1}">
                        <div class="goal-header"><h4 class="goal-title">${g.title}</h4><span class="badge approved">${g.achStatus || 'Pending Update'}</span></div>
                        <div class="goal-meta"><div class="meta-item"><span class="meta-label">Target</span><span>${g.uom === 'Timeline' ? g.deadline : g.target}</span></div><div class="meta-item"><span class="meta-label">Actual</span><span>${(g.actual !== undefined && g.actual !== null && g.actual !== '') ? g.actual : '-'}</span></div></div>
                        <div class="progress-container"><div class="progress-header"><span>System Progress</span><span>${progress}%</span></div><div class="progress-track"><div class="progress-fill" style="width: ${progress}%"></div></div></div>
                        <div class="form-group" style="margin-top: 1rem;"><label style="font-size:0.75rem;">Manager Check-in Note</label>
                        <textarea id="checkin_${g.id}" rows="2" class="inline-edit" ${isPhase1?'disabled':''}>${g.checkinComment || ''}</textarea></div>
                        <button class="btn-primary btn-save-checkin" style="width: 100%; justify-content: center;" data-id="${g.id}" ${isPhase1?'disabled':''}>Save Check-in</button>
                    </div>`;
            });
            html += `</div>`; mgrContent.innerHTML = html;

            document.querySelectorAll('.btn-save-checkin').forEach(btn => btn.addEventListener('click', (e) => {
                const id = e.target.dataset.id; const goal = db.goals.find(g => g.id === id);
                goal.checkinComment = document.getElementById(`checkin_${id}`).value;
                saveDb(); logAudit('Check-in Updated', `Manager updated check-in for ${id}`, 'Manager'); showToast('Check-in saved.', 'success');
            }));
        };

        // CSV Export
        document.getElementById('btnExportCSV').addEventListener('click', () => {
            const goals = db.goals.filter(g => g.status === 'Approved');
            if (goals.length === 0) return showToast('No approved goals to export.', 'error');
            
            let csvContent = "data:text/csv;charset=utf-8,Employee ID,Thrust Area,Title,UoM,Weightage,Target,Actual,Progress Score\n";
            goals.forEach(g => {
                const row = `emp_1,${g.thrustArea},"${g.title}",${g.uom},${g.weightage},${g.uom==='Timeline'?g.deadline:g.target},${g.actual||0},${calculateProgress(g)}%`;
                csvContent += row + "\n";
            });
            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a"); link.setAttribute("href", encodedUri); link.setAttribute("download", "achievement_report.csv");
            document.body.appendChild(link); link.click(); document.body.removeChild(link);
            logAudit('Report Exported', 'Manager exported CSV', 'Manager');
            showToast('CSV Exported Successfully!', 'success');
        });

        renderApprovals(); 
    };

    // ==========================================
    // ADMIN VIEW
    // ==========================================
    const initAdminView = () => {
        const renderAuditLogs = () => {
            const container = document.getElementById('auditLogContainer');
            if (db.audit.length === 0) return container.innerHTML = '<p class="text-muted">No logs.</p>';
            container.innerHTML = db.audit.map(log => `
                <div class="log-entry"><div style="display:flex; justify-content: space-between;"><strong>${log.action} <span class="badge">${log.user}</span></strong><span class="log-time">${new Date(log.timestamp).toLocaleString()}</span></div><p style="margin-top:0.25rem; color:var(--text-muted);">${log.details}</p></div>
            `).join('');
        };
        renderAuditLogs();

        // Cycle Mgmt
        document.querySelectorAll('.btn-cycle').forEach(btn => {
            if(btn.dataset.cycle === db.activeCycle) {
                document.querySelectorAll('.btn-cycle').forEach(b => b.classList.remove('active')); btn.classList.add('active');
            }
            btn.addEventListener('click', (e) => {
                document.querySelectorAll('.btn-cycle').forEach(b => b.classList.remove('active')); e.target.classList.add('active');
                db.activeCycle = e.target.dataset.cycle; saveDb();
                logAudit('Cycle Changed', `Admin changed cycle to ${getCycleName(db.activeCycle)}`, 'Admin');
                showToast(`Cycle set to ${getCycleName(db.activeCycle)}`, 'success');
                renderAuditLogs();
            });
        });

        let pieChartInst = null;
        let barChartInst = null;

        const renderAnalytics = () => {
            Chart.defaults.color = '#475569';
            Chart.defaults.borderColor = 'rgba(0,0,0,0.05)';

            const areas = { 'Financial':0, 'Customer':0, 'Process':0, 'Learning':0 };
            db.goals.forEach(g => { if(areas[g.thrustArea] !== undefined) areas[g.thrustArea]++; });
            
            const pieCtx = document.getElementById('pieChartCanvas');
            if (pieCtx) {
                if(pieChartInst) pieChartInst.destroy();
                pieChartInst = new Chart(pieCtx, {
                    type: 'doughnut',
                    data: {
                        labels: Object.keys(areas),
                        datasets: [{
                            data: Object.values(areas),
                            backgroundColor: ['#3b82f6', '#10b981', '#8b5cf6', '#f59e0b'],
                            borderWidth: 0,
                            hoverOffset: 10
                        }]
                    },
                    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'right' } } }
                });
            }

            // Manager effectiveness logic
            let approvedGoals = 0; let pendingGoals = 0;
            db.goals.forEach(g => {
                if (['Approved'].includes(g.status)) approvedGoals++;
                else if (['Submitted'].includes(g.status)) pendingGoals++;
            });

            // If there are no goals, fake some data so the chart isn't empty on first load
            const hasData = (approvedGoals + pendingGoals) > 0;
            const janeScore = hasData ? Math.min((approvedGoals / (approvedGoals + pendingGoals)) * 100 + 20, 100) : 85;

            const barCtx = document.getElementById('barChartCanvas');
            if (barCtx) {
                if(barChartInst) barChartInst.destroy();
                barChartInst = new Chart(barCtx, {
                    type: 'bar',
                    data: {
                        labels: ['Jane Smith (Mgr A)', 'Bob (Mgr B)', 'Alice (Mgr C)'],
                        datasets: [{
                            label: 'Completion Rate %',
                            data: [janeScore, 40, 95], // Mocking other managers, tying Jane to actual data roughly
                            backgroundColor: ['#3b82f6', '#f59e0b', '#10b981'],
                            borderRadius: 4
                        }]
                    },
                    options: {
                        responsive: true, maintainAspectRatio: false,
                        scales: { y: { beginAtZero: true, max: 100 } }
                    }
                });
            }
        };
        renderAnalytics();

        // Escalation
        document.getElementById('btnRunEscalation').addEventListener('click', () => {
            let pendingCount = 0;
            let zeroProgressCount = 0;

            db.goals.forEach(g => {
                if (g.status === 'Submitted') pendingCount++;
                if (g.status === 'Approved' && (!g.actual || g.actual == 0) && g.uom !== 'Timeline') zeroProgressCount++;
            });

            logAudit('Escalation Engine', 'Admin manually triggered rule evaluation', 'System');
            
            if (pendingCount > 0) {
                triggerWebhook('MS Teams', `Alert to Manager Jane Smith: You have ${pendingCount} pending goal(s) awaiting approval in your queue.`);
            }
            if (zeroProgressCount > 0) {
                triggerWebhook('Outlook Email', `To: John Doe. Subject: Action Required. You have ${zeroProgressCount} goal(s) with 0 progress. Please update your check-in.`);
            }
            
            if (pendingCount === 0 && zeroProgressCount === 0) {
                showToast('Escalation engine ran. No warnings needed!', 'success');
            } else {
                showToast('Escalation engine run complete. Notifications fired.', 'warning');
            }
            renderAuditLogs();
        });

        // Factory Reset
        document.getElementById('btnFactoryReset')?.addEventListener('click', () => {
            if(confirm("Are you SURE you want to wipe the current database? This cannot be undone.")) {
                db = { goals: [], audit: [], webhooks: [], activeCycle: 'phase1' };
                saveDb();
                showToast("Database has been reset to factory defaults.", "success");
                renderView();
            }
        });

        // Shared Goal
        const modal = document.getElementById('sharedGoalModal');
        document.getElementById('btnPushSharedGoal').addEventListener('click', () => { document.getElementById('sharedGoalForm').reset(); modal.classList.remove('hidden'); });
        document.getElementById('btnCloseSharedModal').addEventListener('click', () => modal.classList.add('hidden')); document.getElementById('btnCancelShared').addEventListener('click', () => modal.classList.add('hidden'));
        document.getElementById('btnSaveShared').addEventListener('click', () => {
            const form = document.getElementById('sharedGoalForm'); if (!form.checkValidity()) return form.reportValidity();
            db.goals.push({ id: generateId(), ownerId: 'emp_1', status: 'Draft', isShared: true, thrustArea: document.getElementById('sgThrustArea').value, title: document.getElementById('sgTitle').value, uom: document.getElementById('sgUom').value, target: document.getElementById('sgTarget').value, weightage: 10, desc: 'KPI pushed by Admin' });
            saveDb(); logAudit('Push Shared Goal', `Admin pushed KPI: ${document.getElementById('sgTitle').value}`, 'Admin');
            modal.classList.add('hidden'); renderAuditLogs(); renderAnalytics(); showToast('Shared goal successfully pushed.', 'success');
        });
    };

    // --- Init Webhooks ---
    renderWebhooks();
});
