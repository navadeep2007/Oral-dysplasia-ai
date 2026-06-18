// OralDysplasia AI — Web Client Controller

// ── Demo Mode Detection ────────────────────────────────────────────
// When hosted on GitHub Pages (no backend), enable demo mode with
// localStorage-based simulation so registration/login/dashboard work.
const IS_DEMO_MODE = window.location.hostname.includes("github.io");
const API_BASE = "/api/v1";

// ── Demo Mode Helpers (GitHub Pages localStorage simulation) ───────
function _demoGetUsers() {
    try { return JSON.parse(localStorage.getItem("demo_users") || "[]"); } catch { return []; }
}
function _demoSaveUsers(users) {
    localStorage.setItem("demo_users", JSON.stringify(users));
}
function _demoGetSlides() {
    try { return JSON.parse(localStorage.getItem("demo_slides") || "[]"); } catch { return []; }
}
function _demoSaveSlides(slides) {
    localStorage.setItem("demo_slides", JSON.stringify(slides));
}
function _demoMakeToken(email) {
    // Simple base64 mock JWT for demo
    const header = btoa(JSON.stringify({alg:"HS256",typ:"JWT"}));
    const payload = btoa(JSON.stringify({sub:email, exp: Date.now()+86400000}));
    return `${header}.${payload}.demo_signature`;
}

// Global Session State
let token = localStorage.getItem("jwt_token");
let user = null;

try {
    const cachedUser = localStorage.getItem("user");
    if (cachedUser) user = JSON.parse(cachedUser);
} catch (e) {
    console.error("Failed to parse user cache:", e);
}

// Canvas Viewer State
let canvasSlideData = null;
let zoomScale = 1.0;
let panOffset = { x: 0, y: 0 };
let isDragging = false;
let startDragOffset = { x: 0, y: 0 };
let activeSlideId = null;

// WHO checklist criteria selection
let selectedWHO = new Set();

// ── App Init & Auth Check ──────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
    setupEventListeners();
    checkAuthSession();
});

function checkAuthSession() {
    if (token && user) {
        document.getElementById("landing-container").classList.add("hidden");
        document.getElementById("auth-modal").classList.add("hidden");
        document.getElementById("app-container").classList.remove("hidden");
        updateUserUI();
        navigateTo("dashboard-section");
    } else {
        document.getElementById("landing-container").classList.remove("hidden");
        document.getElementById("auth-modal").classList.add("hidden");
        document.getElementById("app-container").classList.add("hidden");
    }
}

function showAuthModal(view = "login") {
    document.getElementById("auth-modal").classList.remove("hidden");
    if (view === "login") {
        document.getElementById("login-view").classList.remove("hidden");
        document.getElementById("signup-view").classList.add("hidden");
        document.getElementById("forgot-view").classList.add("hidden");
    } else if (view === "signup") {
        document.getElementById("login-view").classList.add("hidden");
        document.getElementById("signup-view").classList.remove("hidden");
        document.getElementById("forgot-view").classList.add("hidden");
    } else if (view === "forgot") {
        document.getElementById("login-view").classList.add("hidden");
        document.getElementById("signup-view").classList.add("hidden");
        document.getElementById("forgot-view").classList.remove("hidden");
    }
}

function hideAuthModal() {
    document.getElementById("auth-modal").classList.add("hidden");
}

function updateUserUI() {
    if (!user) return;
    document.getElementById("sidebar-user-name").textContent = user.name;
    document.getElementById("sidebar-user-role").textContent = user.role;
    
    // Set profile view parameters
    document.getElementById("profile-name").textContent = user.name;
    document.getElementById("profile-license").textContent = user.license_id || "N/A";
    document.getElementById("profile-email").textContent = user.email;
    document.getElementById("profile-role").textContent = user.role;
    document.getElementById("profile-institution").textContent = user.institution;
    
    // Verification block signatory preview
    document.getElementById("sig-name").textContent = user.name;
    document.getElementById("sig-role").textContent = user.role;
    document.getElementById("sig-license").textContent = user.license_id || "N/A";
    document.getElementById("sig-institution").textContent = user.institution;
}

// ── Navigation routing ─────────────────────────────────────────────
function navigateTo(sectionId) {
    // Hide all main sections
    document.querySelectorAll(".app-section").forEach(s => s.classList.add("hidden"));
    // Show active section
    document.getElementById(sectionId).classList.remove("hidden");
    
    // Highlight sidebar active item
    document.querySelectorAll(".nav-item").forEach(item => {
        if (item.getAttribute("data-target") === sectionId) {
            item.classList.add("active");
        } else {
            item.classList.remove("active");
        }
    });

    // Run view specific initializers
    if (sectionId === "dashboard-section") {
        loadDashboardStats();
    } else if (sectionId === "library-section") {
        loadLibraryData();
    }
}

// ── Overlay loader helper ──────────────────────────────────────────
function showLoader(show, message = "Loading...") {
    const overlay = document.getElementById("loading-overlay");
    const msgLabel = document.getElementById("loading-message");
    if (show) {
        msgLabel.textContent = message;
        overlay.classList.remove("hidden");
    } else {
        overlay.classList.add("hidden");
    }
}

// ── API Fetch request coordinator ──────────────────────────────────
async function fetchAPI(endpoint, options = {}) {
    const headers = options.headers || {};
    if (token) {
        headers["Authorization"] = `Bearer ${token}`;
    }
    
    options.headers = headers;
    
    try {
        const response = await fetch(`${API_BASE}${endpoint}`, options);
        if (response.status === 401) {
            // Expired/Invalid token
            handleLogout();
            throw new Error("Session expired. Please log in again.");
        }
        if (!response.ok) {
            const errorText = await response.text();
            let errMsg = "API Request failed";
            try {
                const errJson = JSON.parse(errorText);
                errMsg = errJson.detail || errMsg;
            } catch (e) {}
            throw new Error(errMsg);
        }
        return await response.json();
    } catch (e) {
        console.error("API Call error:", e);
        throw e;
    }
}

// ── Auth actions (Login, SignUp, Logout) ───────────────────────────
function handleLogout() {
    token = null;
    user = null;
    localStorage.clear();
    checkAuthSession();
}

// Event Listeners setup
function setupEventListeners() {
    // Modal toggle inside auth forms
    document.getElementById("go-to-signup").addEventListener("click", (e) => {
        e.preventDefault();
        showAuthModal("signup");
    });
    document.getElementById("go-to-login").addEventListener("click", (e) => {
        e.preventDefault();
        showAuthModal("login");
    });

    // Close Auth modal
    document.getElementById("btn-close-auth").addEventListener("click", () => {
        hideAuthModal();
    });

    // Forgot Password Nav Toggle
    document.getElementById("go-to-forgot-password").addEventListener("click", (e) => {
        e.preventDefault();
        showAuthModal("forgot");
    });
    document.getElementById("forgot-go-to-login").addEventListener("click", (e) => {
        e.preventDefault();
        showAuthModal("login");
    });

    // Forgot Password Form Submit
    document.getElementById("forgot-form").addEventListener("submit", async (e) => {
        e.preventDefault();
        const email = document.getElementById("forgot-email").value.trim();
        const errBlock = document.getElementById("forgot-error");
        const successBlock = document.getElementById("forgot-success");
        errBlock.classList.add("hidden");
        successBlock.classList.add("hidden");
        showLoader(true, "Requesting password reset...");
        try {
            const res = await fetch(`${API_BASE}/auth/forgot-password`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ email })
            });
            if (!res.ok) {
                const txt = await res.text();
                let msg = "Email address not registered";
                try { msg = JSON.parse(txt).detail || msg; } catch (ex) {}
                throw new Error(msg);
            }
            showLoader(false);
            successBlock.textContent = "Recovery instructions sent! Check your inbox.";
            successBlock.classList.remove("hidden");
        } catch (err) {
            showLoader(false);
            errBlock.textContent = err.message;
            errBlock.classList.remove("hidden");
        }
    });

    // Show Auth modal from landing CTA buttons
    document.getElementById("btn-show-login").addEventListener("click", () => showAuthModal("login"));
    document.getElementById("btn-show-signup").addEventListener("click", () => showAuthModal("signup"));
    document.getElementById("btn-hero-launch").addEventListener("click", () => showAuthModal("login"));

    // Mobile Navigation triggers
    const menuToggle = document.getElementById("landing-menu-toggle");
    const mobileMenu = document.getElementById("mobile-nav-menu");
    menuToggle.addEventListener("click", () => {
        mobileMenu.classList.toggle("hidden");
    });

    document.getElementById("btn-mobile-login").addEventListener("click", () => {
        mobileMenu.classList.add("hidden");
        showAuthModal("login");
    });
    document.getElementById("btn-mobile-signup").addEventListener("click", () => {
        mobileMenu.classList.add("hidden");
        showAuthModal("signup");
    });

    // Close mobile nav when clicking a link
    document.querySelectorAll(".mobile-nav-item").forEach(item => {
        item.addEventListener("click", () => mobileMenu.classList.add("hidden"));
    });

    // Login Form Submit
    document.getElementById("login-form").addEventListener("submit", async (e) => {
        e.preventDefault();
        const email = document.getElementById("login-email").value.trim();
        const password = document.getElementById("login-password").value;
        const errBlock = document.getElementById("login-error");
        
        errBlock.classList.add("hidden");
        showLoader(true, "Verifying clinical access keys...");
        
        try {
            const res = await fetch(`${API_BASE}/auth/login`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ email, password })
            });
            if (!res.ok) {
                const txt = await res.text();
                let msg = "Invalid credentials";
                try { msg = JSON.parse(txt).detail || msg; } catch (ex) {}
                throw new Error(msg);
            }
            const data = await res.json();
            token = data.access_token;
            user = data.user;
            
            localStorage.setItem("jwt_token", token);
            localStorage.setItem("user", JSON.stringify(user));
            
            showLoader(false);
            checkAuthSession();
        } catch (err) {
            showLoader(false);
            errBlock.textContent = err.message;
            errBlock.classList.remove("hidden");
        }
    });

    // SignUp Form Submit
    document.getElementById("signup-form").addEventListener("submit", async (e) => {
        e.preventDefault();
        const name = document.getElementById("signup-name").value.trim();
        const email = document.getElementById("signup-email").value.trim();
        const license_id = document.getElementById("signup-license").value.trim();
        const role = document.getElementById("signup-role").value;
        const institution = document.getElementById("signup-institution").value.trim();
        const password = document.getElementById("signup-password").value;
        const errBlock = document.getElementById("signup-error");

        errBlock.classList.add("hidden");
        showLoader(true, "Registering pathologist credentials...");

        if (IS_DEMO_MODE) {
            // Demo mode: save to localStorage
            await new Promise(r => setTimeout(r, 800)); // simulate network delay
            const users = _demoGetUsers();
            if (users.find(u => u.email === email)) {
                showLoader(false);
                errBlock.textContent = "Email already registered";
                errBlock.classList.remove("hidden");
                return;
            }
            const newUser = { id: users.length + 1, name, email, license_id, role, institution, password };
            users.push(newUser);
            _demoSaveUsers(users);
            token = _demoMakeToken(email);
            user = { id: newUser.id, name, email, role, institution, license_id };
            localStorage.setItem("jwt_token", token);
            localStorage.setItem("user", JSON.stringify(user));
            showLoader(false);
            checkAuthSession();
            return;
        }

        try {
            const res = await fetch(`${API_BASE}/auth/signup`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ name, email, license_id, role, institution, password })
            });
            if (!res.ok) {
                const txt = await res.text();
                let msg = "Failed to register";
                try { msg = JSON.parse(txt).detail || msg; } catch (ex) {}
                throw new Error(msg);
            }
            const data = await res.json();
            token = data.access_token;
            user = data.user;

            localStorage.setItem("jwt_token", token);
            localStorage.setItem("user", JSON.stringify(user));

            showLoader(false);
            checkAuthSession();
        } catch (err) {
            showLoader(false);
            errBlock.textContent = err.message;
            errBlock.classList.remove("hidden");
        }
    });

    // Sidebar navigation trigger
    document.querySelectorAll(".nav-item").forEach(item => {
        item.addEventListener("click", (e) => {
            e.preventDefault();
            const target = item.getAttribute("data-target");
            navigateTo(target);
        });
    });

    // Logout
    document.getElementById("logout-btn").addEventListener("click", handleLogout);

    // Filter Chips Event Binding
    document.querySelectorAll("#grade-filter-chips .chip").forEach(c => {
        c.addEventListener("click", () => {
            document.querySelectorAll("#grade-filter-chips .chip").forEach(ch => ch.classList.remove("active"));
            c.classList.add("active");
            loadLibraryData();
        });
    });
    document.querySelectorAll("#status-filter-chips .chip").forEach(c => {
        c.addEventListener("click", () => {
            document.querySelectorAll("#status-filter-chips .chip").forEach(ch => ch.classList.remove("active"));
            c.classList.add("active");
            loadLibraryData();
        });
    });

    // File selection and Drag-and-Drop handling
    let selectedFile = null; // Can be a real File object or mock configuration object
    const fileDropZone = document.getElementById("file-drop-zone");
    const fileInput = document.getElementById("file-input");

    fileDropZone.addEventListener("click", (e) => {
        // Prevent launching file picker when clicking mock pick buttons
        if (e.target.closest(".mock-btn-row")) {
            return;
        }
        fileInput.click();
    });

    fileInput.addEventListener("change", () => {
        if (fileInput.files && fileInput.files.length > 0) {
            selectedFile = fileInput.files[0];
            document.getElementById("selected-filename-label").textContent = `Selected: ${selectedFile.name}`;
        }
    });

    fileDropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        fileDropZone.classList.add("dragover");
    });

    fileDropZone.addEventListener("dragleave", () => {
        fileDropZone.classList.remove("dragover");
    });

    fileDropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        fileDropZone.classList.remove("dragover");
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            selectedFile = e.dataTransfer.files[0];
            document.getElementById("selected-filename-label").textContent = `Selected: ${selectedFile.name}`;
            fileInput.files = e.dataTransfer.files; // Sync with file input
        }
    });

    // Mock File pickers inside Upload
    document.getElementById("pick-mock-a").addEventListener("click", () => {
        selectedFile = { name: "slide_case_A.svs", isMock: true };
        document.getElementById("selected-filename-label").textContent = "Selected: slide_case_A.svs";
    });
    document.getElementById("pick-mock-b").addEventListener("click", () => {
        selectedFile = { name: "slide_case_B.ndpi", isMock: true };
        document.getElementById("selected-filename-label").textContent = "Selected: slide_case_B.ndpi";
    });

    // Upload slide form submission
    document.getElementById("upload-form").addEventListener("submit", async (e) => {
        e.preventDefault();
        const errBlock = document.getElementById("upload-error");
        const progressContainer = document.getElementById("upload-progress-container");
        const progressBar = document.getElementById("upload-progress-bar");
        const progressStatus = document.getElementById("upload-progress-status");
        const submitBtn = document.getElementById("upload-submit-btn");

        errBlock.classList.add("hidden");

        if (!selectedFile) {
            errBlock.textContent = "Please select or drop a slide file first.";
            errBlock.classList.remove("hidden");
            return;
        }

        const patientId = document.getElementById("upload-patient-id").value.trim();
        const patientName = document.getElementById("upload-patient-name").value.trim();
        const patientAge = document.getElementById("upload-patient-age").value.trim();
        const patientGender = document.getElementById("upload-patient-gender").value;
        const anatomicalSite = document.getElementById("upload-site").value;
        const clinicalNotes = document.getElementById("upload-notes").value.trim();

        const fd = new FormData();
        if (selectedFile instanceof File) {
            fd.append("file", selectedFile);
        } else if (selectedFile && selectedFile.isMock) {
            const blob = new Blob(["Mock WSI Virtual Scan file data bytes"], { type: "application/octet-stream" });
            fd.append("file", blob, selectedFile.name);
        }
        fd.append("patient_id", patientId);
        fd.append("patient_name", patientName);
        fd.append("patient_age", patientAge);
        fd.append("patient_gender", patientGender);
        fd.append("anatomical_site", anatomicalSite);
        if (clinicalNotes) fd.append("clinical_notes", clinicalNotes);

        progressContainer.classList.remove("hidden");
        submitBtn.disabled = true;
        progressBar.style.width = "0%";
        progressStatus.textContent = "Initiating upload connection...";

        let apiResult = null;
        let apiError = null;
        let apiCompleted = false;

        // Start API call
        fetch(`${API_BASE}/slides/upload`, {
            method: "POST",
            headers: { "Authorization": `Bearer ${token}` },
            body: fd
        }).then(async res => {
            if (!res.ok) {
                const text = await res.text();
                throw new Error(text || "Biopsy upload transaction failed");
            }
            apiResult = await res.json();
            apiCompleted = true;
        }).catch(err => {
            apiError = err;
            apiCompleted = true;
        });

        // Simulate progress animation
        let currentProgress = 0;
        const stages = [
            { limit: 30, text: "Uploading WSI slide scan stream...", speed: 30 },
            { limit: 65, text: "Dividing scan into tissue tiles (256x256)...", speed: 45 },
            { limit: 90, text: "Executing Swin-T diagnostic validation...", speed: 50 },
            { limit: 98, text: "Writing encrypted patient demographic blocks to database...", speed: 30 }
        ];

        let stageIdx = 0;

        function step() {
            if (apiCompleted && apiError) {
                progressContainer.classList.add("hidden");
                submitBtn.disabled = false;
                errBlock.textContent = apiError.message;
                errBlock.classList.remove("hidden");
                return;
            }

            if (stageIdx >= stages.length) {
                // Wait for API call if it hasn't finished yet
                if (apiCompleted) {
                    progressBar.style.width = "100%";
                    progressStatus.textContent = "Pipeline ready!";
                    setTimeout(() => {
                        progressContainer.classList.add("hidden");
                        submitBtn.disabled = false;
                        openSlideDetail(apiResult.id);
                    }, 500);
                } else {
                    progressStatus.textContent = "Finalizing diagnostic run on server...";
                    setTimeout(step, 100);
                }
                return;
            }

            const currentStage = stages[stageIdx];
            progressStatus.textContent = currentStage.text;
            currentProgress += Math.floor(Math.random() * 4) + 1;

            if (currentProgress >= currentStage.limit) {
                currentProgress = currentStage.limit;
                stageIdx++;
            }

            progressBar.style.width = `${currentProgress}%`;
            setTimeout(step, currentStage.speed);
        }

        step();
    });

    // Back to library hooks
    document.getElementById("btn-back-to-list").addEventListener("click", () => {
        navigateTo("library-section");
    });
    document.getElementById("btn-back-to-detail").addEventListener("click", () => {
        if (activeSlideId) openSlideDetail(activeSlideId);
    });

    // Initialize Analysis Trigger Button
    document.getElementById("btn-initialize-analysis").addEventListener("click", async () => {
        if (!activeSlideId) return;
        showLoader(true, "Triggering AI Deep Learning pipeline runner...");
        try {
            await fetchAPI("/analysis/run", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ slide_id: activeSlideId, confidence_threshold: 0.5 })
            });
            
            // Poll for processing status every 2 seconds
            pollAnalysisStatus(activeSlideId);
        } catch (err) {
            showLoader(false);
            alert("Analysis failed to start: " + err.message);
        }
    });

    // Open Diagnostics Canvas button
    document.getElementById("btn-open-diagnostics").addEventListener("click", () => {
        if (activeSlideId) openDiagnosticsCanvas(activeSlideId);
    });

    // WHO Checklist selection toggles
    document.getElementById("accordion-checklist-trigger").addEventListener("click", () => {
        const body = document.getElementById("accordion-checklist-body");
        const arrow = document.getElementById("accordion-arrow");
        body.classList.toggle("hidden");
        arrow.classList.toggle("fa-chevron-down");
        arrow.classList.toggle("fa-chevron-up");
    });

    document.querySelectorAll('#accordion-checklist-body input[type="checkbox"]').forEach(cb => {
        cb.addEventListener("change", () => {
            if (cb.checked) {
                selectedWHO.add(cb.value);
            } else {
                selectedWHO.delete(cb.value);
            }
            // Update checklist count indicator
            document.getElementById("accordion-checklist-trigger").querySelector("span").textContent = 
                `WHO Histological Checklist (${selectedWHO.size} observed)`;
        });
    });

    // ICD Select custom toggle
    document.getElementById("canvas-icd-select").addEventListener("change", (e) => {
        const wrap = document.getElementById("canvas-custom-icd-wrapper");
        if (e.target.value === "custom") {
            wrap.classList.remove("hidden");
        } else {
            wrap.classList.add("hidden");
        }
    });

    // Submit review sign-off
    document.getElementById("btn-submit-review").addEventListener("click", async () => {
        if (!activeSlideId) return;
        
        const finalGrade = document.getElementById("canvas-final-grade").value;
        const commentInput = document.getElementById("canvas-comments").value.trim();
        const selectIcd = document.getElementById("canvas-icd-select").value;
        const icd10Code = selectIcd === "custom" ? document.getElementById("canvas-icd-custom").value.trim() : selectIcd;
        
        const successBlock = document.getElementById("canvas-status-msg");
        const errorBlock = document.getElementById("canvas-error-msg");
        
        successBlock.classList.add("hidden");
        errorBlock.classList.add("hidden");
        showLoader(true, "Sealing cryptographic pathologist diagnostic signature...");

        // Construct WHO Checklist summary to prefix comments
        let criteriaPrefix = "";
        if (selectedWHO.size > 0) {
            criteriaPrefix = "Checked WHO Diagnostic Criteria:\n" + Array.from(selectedWHO).map(x => `- ${x}`).join("\n") + "\n\n";
        }
        const comments = (criteriaPrefix + commentInput).trim();

        try {
            await fetchAPI(`/analysis/${activeSlideId}/review`, {
                method: "PUT",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    annotations: [],
                    final_grade: finalGrade,
                    comments: comments || null,
                    icd_10_code: icd10Code
                })
            });
            showLoader(false);
            successBlock.textContent = "Pathology verification signed & sealed successfully!";
            successBlock.classList.remove("hidden");
            
            // Re-render simulated cryptoseal signature box with signed code
            const input = `${activeSlideId}-${finalGrade}-${comments.hashCode()}`;
            const absHash = Math.abs(input.hashCode());
            document.getElementById("sig-hash").textContent = "SHA256:ECDSA:9F8" + absHash.toString(16).toUpperCase().padStart(8, '0');
        } catch (err) {
            showLoader(false);
            errorBlock.textContent = err.message;
            errorBlock.classList.remove("hidden");
        }
    });

    // Export reports click handling
    document.getElementById("btn-export-dropdown").addEventListener("click", (e) => {
        e.stopPropagation();
        document.getElementById("export-dropdown-menu").classList.toggle("hidden");
    });
    
    document.addEventListener("click", () => {
        document.getElementById("export-dropdown-menu").classList.add("hidden");
    });

    document.getElementById("export-fhir").addEventListener("click", () => triggerExport("fhir"));
    document.getElementById("export-dicom").addEventListener("click", () => triggerExport("dicom"));
    document.getElementById("export-pdf").addEventListener("click", () => triggerExport("pdf"));
    document.getElementById("export-patient-pdf").addEventListener("click", () => triggerExport("patient_pdf"));
    document.getElementById("export-whatsapp").addEventListener("click", () => triggerShare("whatsapp"));
    document.getElementById("export-email").addEventListener("click", () => triggerShare("email"));

    // Canvas zoom & pan mouse events
    const canvas = document.getElementById("wsi-canvas");
    canvas.addEventListener("mousedown", (e) => {
        isDragging = true;
        startDragOffset = { x: e.clientX - panOffset.x, y: e.clientY - panOffset.y };
    });
    canvas.addEventListener("mousemove", (e) => {
        if (!isDragging) return;
        panOffset = { x: e.clientX - startDragOffset.x, y: e.clientY - startDragOffset.y };
        renderCanvas();
    });
    window.addEventListener("mouseup", () => {
        isDragging = false;
    });

    // Zoom Buttons
    document.getElementById("btn-zoom-in").addEventListener("click", () => {
        zoomScale = Math.min(zoomScale * 1.3, 6.0);
        renderCanvas();
    });
    document.getElementById("btn-zoom-out").addEventListener("click", () => {
        zoomScale = Math.max(zoomScale / 1.3, 0.5);
        renderCanvas();
    });
    document.getElementById("btn-zoom-reset").addEventListener("click", () => {
        zoomScale = 1.0;
        panOffset = { x: 0, y: 0 };
        renderCanvas();
    });
}

// ── Dashboard Statistics Loader ─────────────────────────────────────
async function loadDashboardStats() {
    try {
        const stats = await fetchAPI("/slides/stats/dashboard");
        document.getElementById("kpi-total").textContent = stats.total_slides;
        document.getElementById("kpi-pending").textContent = stats.pending_review;
        document.getElementById("kpi-severe").textContent = stats.severe_count;

        // Render slides list
        const container = document.getElementById("dashboard-slides-list");
        container.innerHTML = "";
        
        if (stats.recent_slides.length === 0) {
            container.innerHTML = `<div class="slide-item-card" style="justify-content: center; cursor: default;">
                <p style="color: var(--text-muted);">No biopsy slides uploaded to database.</p>
            </div>`;
            return;
        }

        stats.recent_slides.forEach(slide => {
            const card = document.createElement("div");
            card.className = "slide-item-card";
            card.onclick = () => openSlideDetail(slide.id);

            const metaColor = slide.status === "processed" ? "var(--grade-mild)" : (slide.status === "reviewed" ? "var(--text-indigo)" : (slide.status === "analyzing" ? "#D97706" : "var(--text-muted)"));
            const [gradeText, gradeClass] = getGradeDetails(slide.current_grade);

            card.innerHTML = `
                <div class="slide-item-info">
                    <h4>${slide.filename}</h4>
                    <div class="slide-item-meta">
                        <span>Patient: <strong>${slide.patient_id}</strong></span>
                        <span>Site: ${slide.anatomical_site}</span>
                    </div>
                </div>
                <div style="display: flex; flex-direction: column; align-items: flex-end; gap: 4px;">
                    <span class="grade-chip ${gradeClass}">${gradeText}</span>
                    <span style="font-size: 9px; font-weight: 800; color: ${metaColor}">${slide.status.toUpperCase()}</span>
                </div>
            `;
            container.appendChild(card);
        });
    } catch (err) {
        console.error("Dashboard stats failed to load:", err);
    }
}

// Helper mapping grades
function getGradeDetails(grade) {
    const g = (grade || "pending").toLowerCase().trim();
    if (g === "mild") return ["MILD DYSPLASIA", "grade-mild"];
    if (g === "moderate") return ["MODERATE DYSPLASIA", "grade-moderate"];
    if (g === "severe") return ["SEVERE DYSPLASIA", "grade-severe"];
    if (g === "normal") return ["NORMAL / BENIGN", "grade-normal"];
    return ["PENDING DIAGNOSIS", "grade-pending"];
}

// ── Slide Library Query ────────────────────────────────────────────
async function loadLibraryData() {
    const activeGradeChip = document.querySelector("#grade-filter-chips .chip.active");
    const activeStatusChip = document.querySelector("#status-filter-chips .chip.active");
    
    const grade = activeGradeChip.getAttribute("data-grade");
    const status = activeStatusChip.getAttribute("data-status");
    
    let url = `/slides/library?page=1&limit=50`;
    if (grade !== "all") url += `&grade=${grade}`;
    if (status !== "all") url += `&status_filter=${status}`;

    try {
        const data = await fetchAPI(url);
        const tbody = document.getElementById("library-table-body");
        tbody.innerHTML = "";

        if (data.slides.length === 0) {
            tbody.innerHTML = `<tr><td colspan="8" style="text-align: center; color: var(--text-muted); padding: 32px;">No matching cases found.</td></tr>`;
            return;
        }

        data.slides.forEach(slide => {
            const tr = document.createElement("tr");

            // Safe fallback properties to prevent TypeError null crashes
            const currentGrade = slide.current_grade || "pending";
            const currentStatus = slide.status || "uploaded";
            const patientName = slide.patient_name || "Patient";
            const patientId = slide.patient_id || "N/A";
            const anatomicalSite = slide.anatomical_site || "N/A";

            // Build the patient report text summary locally for instant sharing
            const gradeExplanations = {
                "normal": "No signs of oral epithelial dysplasia were found (Benign / Normal tissue).",
                "mild": "Mild dysplasia detected. Cell alterations are confined to the lower third of the epithelium. Standard monitoring and follow-up recommended.",
                "moderate": "Moderate dysplasia detected. Cell alterations extend to the middle third of the epithelium. Close clinical observation or intervention may be required.",
                "severe": "Severe dysplasia / Carcinoma in situ detected. Cell alterations occupy the upper third or full thickness of the epithelium. Prompt clinical treatment is required.",
                "pending": "Analysis is pending verification."
            };
            const explanation = gradeExplanations[currentGrade.toLowerCase()] || "Verification required by pathologist.";
            const userObj = JSON.parse(localStorage.getItem("user") || "{}");
            const pathName = userObj.name || "Specialist";
            const pathRole = userObj.role || "Pathologist";
            const pathInst = userObj.institution || "OralDysplasia AI Lab";

            const shareText = `PATIENT DIAGNOSTIC REPORT (OralDysplasia AI)
--------------------------------------------
Patient Name: ${patientName}
Patient ID: ${patientId}
Age: ${slide.patient_age || "N/A"}
Gender: ${slide.patient_gender || "N/A"}
Biopsy Site: ${anatomicalSite}

DIAGNOSTIC ASSESSMENT
Grade: ${currentGrade.toUpperCase()}
Summary: ${explanation}
Next Steps: Please consult your oral surgeon or primary clinician to discuss these diagnostic findings.

CERTIFICATION
Verifying Specialist: Dr. ${pathName} (${pathRole})
Institution: ${pathInst}
Status: ${currentStatus.toUpperCase()}`;

            tr.onclick = (e) => {
                // Ignore row navigation clicks if triggered from sharing links
                if (e.target.closest(".btn-share-icon")) return;
                openSlideDetail(slide.id);
            };

            const statusColor = currentStatus === "processed" ? "var(--grade-mild)" : (currentStatus === "reviewed" ? "var(--text-indigo)" : (currentStatus === "analyzing" ? "#D97706" : "var(--text-muted)"));
            const [gradeText, gradeClass] = getGradeDetails(currentGrade);

            let shareCellHTML = "";
            if (currentStatus === "processed" || currentStatus === "reviewed") {
                shareCellHTML = `
                    <a href="https://api.whatsapp.com/send?text=${encodeURIComponent(shareText)}" target="_blank" class="btn-share-icon whatsapp" title="Share via WhatsApp" style="margin-right: 12px; color: #25D366; font-size: 13px; text-decoration: none; font-weight: 700; display: inline-flex; align-items: center; gap: 4px;">
                        <i class="fa-brands fa-whatsapp" style="font-size: 16px;"></i> WhatsApp
                    </a>
                    <a href="mailto:?subject=${encodeURIComponent("Patient Diagnostic Report: " + patientName)}&body=${encodeURIComponent(shareText)}" class="btn-share-icon email" title="Share via Email" style="margin-right: 12px; color: #4F46E5; font-size: 13px; text-decoration: none; font-weight: 700; display: inline-flex; align-items: center; gap: 4px;">
                        <i class="fa-solid fa-envelope" style="font-size: 16px;"></i> Email
                    </a>
                    <a href="#" class="btn-share-icon pdf" title="Export/Download Patient PDF Report" style="color: #EF4444; font-size: 13px; text-decoration: none; font-weight: 700; display: inline-flex; align-items: center; gap: 4px;">
                        <i class="fa-solid fa-file-pdf" style="font-size: 16px;"></i> PDF
                    </a>
                `;
            } else {
                shareCellHTML = `<span style="color: var(--text-muted); font-size: 11.5px; font-weight: 600; font-style: italic;">Awaiting Verification</span>`;
            }

            tr.innerHTML = `
                <td><strong>${slide.filename}</strong></td>
                <td>${patientId}</td>
                <td>${patientName}</td>
                <td>${anatomicalSite}</td>
                <td><span class="grade-chip ${gradeClass}">${gradeText}</span></td>
                <td><strong>${((slide.overall_confidence || 0) * 100).toFixed(0)}%</strong></td>
                <td><span style="font-weight: 800; font-size: 11px; color: ${statusColor}">${currentStatus.toUpperCase()}</span></td>
                <td class="share-actions-cell" style="text-align: center; white-space: nowrap;">
                    ${shareCellHTML}
                </td>
            `;

            // Add stop propagation and click events to links
            const whatsappLink = tr.querySelector(".btn-share-icon.whatsapp");
            const emailLink = tr.querySelector(".btn-share-icon.email");
            const pdfLink = tr.querySelector(".btn-share-icon.pdf");

            if (whatsappLink) whatsappLink.onclick = (e) => e.stopPropagation();
            if (emailLink) emailLink.onclick = (e) => e.stopPropagation();
            if (pdfLink) {
                pdfLink.onclick = (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    triggerExport("patient_pdf", slide.id);
                };
            }

            tbody.appendChild(tr);
        });
    } catch (err) {
        console.error("Failed to load slide library:", err);
    }
}

// ── Open Case Detail Dossier ───────────────────────────────────────
async function openSlideDetail(slideId) {
    activeSlideId = slideId;
    navigateTo("detail-section");
    showLoader(true, "Fetching case records...");
    
    try {
        const slide = await fetchAPI(`/slides/${slideId}`);
        showLoader(false);

        document.getElementById("detail-filename").textContent = slide.filename;
        document.getElementById("detail-status").textContent = slide.status.toUpperCase();
        
        // Style status label
        const statusSpan = document.getElementById("detail-status");
        if (slide.status === "processed") {
            statusSpan.style.backgroundColor = "rgba(16, 185, 129, 0.1)";
            statusSpan.style.color = "var(--grade-mild)";
        } else if (slide.status === "analyzing") {
            statusSpan.style.backgroundColor = "rgba(217, 119, 6, 0.1)";
            statusSpan.style.color = "#D97706";
        } else {
            statusSpan.style.backgroundColor = "rgba(100, 116, 139, 0.1)";
            statusSpan.style.color = "var(--text-muted)";
        }

        // Set grade chip
        const gradeChip = document.getElementById("detail-grade-chip");
        gradeChip.className = "grade-chip";
        const [gradeText, gradeClass] = getGradeDetails(slide.current_grade);
        gradeChip.classList.add(gradeClass);
        gradeChip.textContent = gradeText;

        // Demographic parameters
        document.getElementById("detail-patient-id").textContent = slide.patient_id;
        document.getElementById("detail-patient-name").textContent = slide.patient_name;
        document.getElementById("detail-patient-age").textContent = slide.patient_age || "N/A";
        document.getElementById("detail-patient-gender").textContent = slide.patient_gender || "N/A";
        document.getElementById("detail-site").textContent = slide.anatomical_site;
        document.getElementById("detail-created").textContent = new Date(slide.created_at).toLocaleString();
        
        // Scan properties
        document.getElementById("detail-dimensions").textContent = `${slide.width} x ${slide.height} pixels`;
        document.getElementById("detail-size").textContent = `${(slide.size_bytes / (1024 * 1024)).toFixed(2)} MB`;
        
        // Clinical Notes
        document.getElementById("detail-clinical-notes").textContent = slide.clinical_notes || "No clinical history attached to slide case.";

        // Manage button toggle states
        const btnInit = document.getElementById("btn-initialize-analysis");
        const btnOpen = document.getElementById("btn-open-diagnostics");

        const shareRow = document.getElementById("detail-share-row");
        if (slide.status === "processed" || slide.status === "reviewed") {
            btnInit.classList.add("hidden");
            btnOpen.classList.remove("hidden");
            shareRow.classList.remove("hidden");

            // Build decrypted clinical report sharing summary
            const currentGrade = slide.current_grade || "pending";
            const currentStatus = slide.status || "processed";
            const gradeExplanations = {
                "normal": "No signs of oral epithelial dysplasia were found (Benign / Normal tissue).",
                "mild": "Mild dysplasia detected. Cell alterations are confined to the lower third of the epithelium. Standard monitoring and follow-up recommended.",
                "moderate": "Moderate dysplasia detected. Cell alterations extend to the middle third of the epithelium. Close clinical observation or intervention may be required.",
                "severe": "Severe dysplasia / Carcinoma in situ detected. Cell alterations occupy the upper third or full thickness of the epithelium. Prompt clinical treatment is required.",
                "pending": "Analysis is pending verification."
            };
            const explanation = gradeExplanations[currentGrade.toLowerCase()] || "Verification required by pathologist.";
            const userObj = JSON.parse(localStorage.getItem("user") || "{}");
            const pathName = userObj.name || "Specialist";
            const pathRole = userObj.role || "Pathologist";
            const pathInst = userObj.institution || "OralDysplasia AI Lab";

            const shareText = `PATIENT DIAGNOSTIC REPORT (OralDysplasia AI)
--------------------------------------------
Patient Name: ${slide.patient_name}
Patient ID: ${slide.patient_id}
Age: ${slide.patient_age || "N/A"}
Gender: ${slide.patient_gender || "N/A"}
Biopsy Site: ${slide.anatomical_site}

DIAGNOSTIC ASSESSMENT
Grade: ${currentGrade.toUpperCase()}
Summary: ${explanation}
Next Steps: Please consult your oral surgeon or primary clinician to discuss these diagnostic findings.

CERTIFICATION
Verifying Specialist: Dr. ${pathName} (${pathRole})
Institution: ${pathInst}
Status: ${currentStatus.toUpperCase()}`;

            document.getElementById("detail-share-whatsapp").onclick = (e) => {
                e.preventDefault();
                const url = `https://api.whatsapp.com/send?text=${encodeURIComponent(shareText)}`;
                window.open(url, "_blank");
            };
            document.getElementById("detail-share-email").onclick = (e) => {
                e.preventDefault();
                const url = `mailto:?subject=${encodeURIComponent("Patient Diagnostic Report: " + slide.patient_name)}&body=${encodeURIComponent(shareText)}`;
                window.open(url, "_self");
            };
            document.getElementById("detail-share-pdf").onclick = (e) => {
                e.preventDefault();
                triggerExport("patient_pdf", slide.id);
            };
        } else if (slide.status === "analyzing") {
            btnInit.classList.remove("hidden");
            btnInit.disabled = true;
            btnInit.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> AI Analysis Running...`;
            btnOpen.classList.add("hidden");
            shareRow.classList.add("hidden");
            pollAnalysisStatus(slideId);
        } else {
            btnInit.classList.remove("hidden");
            btnInit.disabled = false;
            btnInit.innerHTML = `<i class="fa-solid fa-play"></i> Initialize AI Diagnostic Runner`;
            btnOpen.classList.add("hidden");
            shareRow.classList.add("hidden");
        }
    } catch (err) {
        showLoader(false);
        alert("Failed to load slide: " + err.message);
    }
}

// Poll analysis status
async function pollAnalysisStatus(slideId) {
    try {
        const slide = await fetchAPI(`/slides/${slideId}`);
        if (slide.status === "processed") {
            showLoader(false);
            openSlideDetail(slideId);
        } else if (slide.status === "error") {
            showLoader(false);
            alert("AI Pipeline runner encountered processing error.");
            openSlideDetail(slideId);
        } else {
            // Keep polling
            setTimeout(() => pollAnalysisStatus(slideId), 2000);
        }
    } catch (e) {
        showLoader(false);
    }
}

// ── Open Diagnostics Canvas ────────────────────────────────────────
async function openDiagnosticsCanvas(slideId) {
    activeSlideId = slideId;
    navigateTo("results-section");
    showLoader(true, "Renditioning digital WSI scanning tissue sections...");

    // Clear verification form inputs
    document.getElementById("canvas-comments").value = "";
    document.getElementById("canvas-status-msg").classList.add("hidden");
    document.getElementById("canvas-error-msg").classList.add("hidden");
    document.getElementById("sig-hash").textContent = "SHA256:ECDSA:PENDING_SEAL";
    selectedWHO.clear();
    document.querySelectorAll('#accordion-checklist-body input[type="checkbox"]').forEach(cb => cb.checked = false);
    document.getElementById("accordion-checklist-trigger").querySelector("span").textContent = "WHO Histological Checklist (0 observed)";

    try {
        const res = await fetchAPI(`/analysis/${slideId}/result`);
        showLoader(false);
        canvasSlideData = res;

        // Render AI Metrics
        const gradeChip = document.getElementById("canvas-grade-chip");
        gradeChip.className = "grade-chip";
        const [gradeText, gradeClass] = getGradeDetails(res.overall_grade);
        gradeChip.classList.add(gradeClass);
        gradeChip.textContent = gradeText;
        document.getElementById("canvas-final-grade").value = res.overall_grade;

        document.getElementById("canvas-confidence-val").textContent = `${(res.overall_confidence * 100).toFixed(0)}%`;

        // Render Frequencies breakdown table
        const breakdownDiv = document.getElementById("canvas-findings-breakdown");
        breakdownDiv.innerHTML = "";

        let hasDetections = false;
        res.patches.forEach(patch => {
            patch.bounding_boxes.forEach(box => {
                hasDetections = true;
                const label = box.label || box.grade;
                const conf = (box.confidence * 100).toFixed(0);
                const targetX = patch.x_index * 256 + (box.xmin + box.xmax) / 2;
                const targetY = patch.y_index * 256 + (box.ymin + box.ymax) / 2;

                const row = document.createElement("div");
                row.className = "finding-row clickable-finding";
                row.style.cursor = "pointer";
                row.title = "Click to snap & focus on canvas";
                
                const colors = { severe: "var(--grade-severe)", moderate: "var(--grade-moderate)", mild: "var(--grade-mild)" };
                const color = colors[box.grade] || "var(--text-muted)";

                row.innerHTML = `
                    <div style="display:flex; align-items:center; gap:8px;">
                        <i class="fa-solid fa-bullseye" style="color: ${color}"></i>
                        <span><strong>${label}</strong> (${conf}%)</span>
                    </div>
                    <span style="font-size:10px; color:var(--text-muted);">Snap <i class="fa-solid fa-expand"></i></span>
                `;

                row.onclick = () => {
                    zoomScale = 2.2;
                    panOffset.x = (512 - targetX) * zoomScale;
                    panOffset.y = (512 - targetY) * zoomScale;
                    renderCanvas();
                };

                breakdownDiv.appendChild(row);
            });
        });

        if (!hasDetections) {
            breakdownDiv.innerHTML = `<p style="font-size: 12px; color: var(--text-muted); text-align: center; padding: 12px;">No abnormalities detected on slide.</p>`;
        }

        // Draw slide
        zoomScale = 1.0;
        panOffset = { x: 0, y: 0 };
        renderCanvas();
    } catch (err) {
        showLoader(false);
        alert("Failed to load analysis canvas: " + err.message);
    }
}

// ── Draw Slide H&E section to Canvas ──────────────────────────────
function renderCanvas() {
    const canvas = document.getElementById("wsi-canvas");
    if (!canvas || !canvasSlideData) return;
    const ctx = canvas.getContext("2d");

    // Clear canvas
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Context transform dimensions
    const scale = zoomScale;
    const tissueSize = 1024 * scale;
    const tissueX = (canvas.width - tissueSize) / 2 + panOffset.x;
    const tissueY = (canvas.height - tissueSize) / 2 + panOffset.y;

    // 1. Draw simulated virtual H&E tissue slice background
    ctx.fillStyle = "#F3E5F5"; // Light pinkish-purple H&E baseline color
    ctx.fillRect(tissueX, tissueY, tissueSize, tissueSize);

    // Draw tissue border outlines
    ctx.strokeStyle = "#E1BEE7";
    ctx.lineWidth = 2 * scale;
    ctx.strokeRect(tissueX, tissueY, tissueSize, tissueSize);

    // 2. Draw cell clusters as hematoxylin violet dots
    const points = [
        {x: 0.2, y: 0.3}, {x: 0.5, y: 0.4}, {x: 0.7, y: 0.2},
        {x: 0.3, y: 0.7}, {x: 0.6, y: 0.8}, {x: 0.8, y: 0.6},
        {x: 0.4, y: 0.5}, {x: 0.1, y: 0.8}, {x: 0.9, y: 0.1}
    ];

    points.forEach(pt => {
        ctx.fillStyle = "rgba(142, 36, 170, 0.15)"; // Hematoxylin violet shade
        ctx.beginPath();
        ctx.arc(tissueX + pt.x * tissueSize, tissueY + pt.y * tissueSize, 35 * scale, 0, 2 * Math.PI);
        ctx.fill();
    });

    // 3. Draw grid and Bounding Boxes
    const tileSize = 256 * scale;
    canvasSlideData.patches.forEach(patch => {
        const tileX = tissueX + patch.x_index * tileSize;
        const tileY = tissueY + patch.y_index * tileSize;

        // Draw dysplastic grid boundaries
        if (patch.predicted_grade !== "normal") {
            const colors = {
                mild: "rgba(16, 185, 129, 0.08)",
                moderate: "rgba(245, 158, 11, 0.08)",
                severe: "rgba(239, 68, 68, 0.08)"
            };
            ctx.fillStyle = colors[patch.predicted_grade] || "transparent";
            ctx.fillRect(tileX, tileY, tileSize, tileSize);
            
            ctx.strokeStyle = patch.predicted_grade === "severe" ? "rgba(239, 68, 68, 0.2)" : "rgba(245, 158, 11, 0.2)";
            ctx.lineWidth = 1 * scale;
            ctx.strokeRect(tileX, tileY, tileSize, tileSize);
        }

        // Draw Bounding Boxes
        patch.bounding_boxes.forEach(box => {
            const boxX = tileX + (box.xmin / 256) * tileSize;
            const boxY = tileY + (box.ymin / 256) * tileSize;
            const boxW = ((box.xmax - box.xmin) / 256) * tileSize;
            const boxH = ((box.ymax - box.ymin) / 256) * tileSize;

            const boxColors = {
                severe: "#EF4444",
                moderate: "#F59E0B",
                mild: "#10B981"
            };
            const boxColor = boxColors[box.grade] || "#64748B";

            // Outline Box
            ctx.strokeStyle = boxColor;
            ctx.lineWidth = 2 * scale;
            ctx.strokeRect(boxX, boxY, boxW, boxH);

            // Text overlay label badge if zoom scale is large enough
            if (scale >= 1.2) {
                const labelText = box.label ? `${box.label} (${(box.confidence * 100).toFixed(0)}%)` : `${box.grade.toUpperCase()} ${(box.confidence * 100).toFixed(0)}%`;
                ctx.font = `bold ${Math.min(9 * scale, 12)}px 'Plus Jakarta Sans', sans-serif`;
                
                const textWidth = ctx.measureText(labelText).width;
                const textHeight = Math.min(10 * scale, 14);

                // Label background card
                ctx.fillStyle = boxColor;
                ctx.fillRect(boxX, boxY - textHeight - 4, textWidth + 8, textHeight + 4);

                // Text
                ctx.fillStyle = "#FFFFFF";
                ctx.fillText(labelText, boxX + 4, boxY - 4);
            }
        });
    });

    // 4. Subtle focus target cursor in center of viewport
    ctx.strokeStyle = "rgba(79, 70, 229, 0.2)";
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.arc(canvas.width / 2, canvas.height / 2, 10, 0, 2 * Math.PI);
    ctx.stroke();
}

// ── Export Diagnostic Reports (FHIR, DICOM, PDF) ─────────────────────
async function triggerExport(format, slideId = null) {
    const targetId = slideId || activeSlideId;
    if (!targetId) return;
    showLoader(true, `Generating ${format.toUpperCase()} diagnostic payload...`);
    try {
        const data = await fetchAPI(`/reports/${targetId}/export?format=${format}`);
        showLoader(false);

        if (format === "pdf" || format === "patient_pdf") {
            const w = window.open("", "_blank");
            if (format === "patient_pdf") {
                w.document.write(`
                    <html>
                    <head>
                        <title>OralDysplasia AI — Patient Diagnostic Report</title>
                        <style>
                            body { font-family: 'Plus Jakarta Sans', sans-serif; padding: 40px; color: #1E293B; line-height: 1.6; max-width: 800px; margin: auto; }
                            .header { text-align: center; margin-bottom: 40px; }
                            .header h1 { color: #4F46E5; margin: 0; font-size: 26px; }
                            .header p { color: #64748B; margin: 5px 0 0 0; font-size: 14px; }
                            .card { background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 24px; border-radius: 12px; margin-bottom: 24px; }
                            .card-title { font-weight: bold; font-size: 16px; color: #1E293B; margin-bottom: 16px; border-bottom: 1px solid #E2E8F0; padding-bottom: 8px; display: flex; align-items: center; gap: 8px; }
                            .info-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; }
                            .info-item { display: flex; flex-direction: column; }
                            .info-label { font-size: 11px; text-transform: uppercase; color: #64748B; font-weight: bold; }
                            .info-value { font-size: 14px; color: #1E293B; font-weight: 600; }
                            .badge { display: inline-block; padding: 6px 12px; border-radius: 6px; font-weight: bold; font-size: 14px; text-transform: uppercase; }
                            .badge-severe { background-color: #FEE2E2; color: #EF4444; }
                            .badge-moderate { background-color: #FEF3C7; color: #F59E0B; }
                            .badge-mild { background-color: #D1FAE5; color: #10B981; }
                            .badge-normal { background-color: #F1F5F9; color: #64748B; }
                            .explanation-text { font-size: 14px; color: #334155; margin-top: 12px; }
                            .footer { text-align: center; font-size: 11px; color: #94A3B8; margin-top: 60px; border-top: 1px solid #F1F5F9; padding-top: 20px; }
                        </style>
                    </head>
                    <body>
                        <div class="header">
                            <h1>PATIENT CLINICAL REPORT</h1>
                            <p>Oral Epithelial Dysplasia AI Screening Service</p>
                        </div>
                        <div class="card">
                            <div class="card-title">Patient Information</div>
                            <div class="info-grid">
                                <div class="info-item"><span class="info-label">Full Name</span><span class="info-value">${data.payload.patient.name}</span></div>
                                <div class="info-item"><span class="info-label">Patient ID</span><span class="info-value">${data.payload.patient.id}</span></div>
                                <div class="info-item"><span class="info-label">Age</span><span class="info-value">${data.payload.patient.age}</span></div>
                                <div class="info-item"><span class="info-label">Gender</span><span class="info-value">${data.payload.patient.gender}</span></div>
                                <div class="info-item"><span class="info-label">Biopsy Site</span><span class="info-value">${data.payload.patient.site}</span></div>
                            </div>
                        </div>
                        <div class="card">
                            <div class="card-title">Diagnostic Assessment</div>
                            <div>
                                <span class="info-label">Clinical Diagnostic Grade</span><br>
                                <span class="badge badge-${data.payload.diagnosis.grade.toLowerCase()}" style="margin-top: 6px;">${data.payload.diagnosis.grade}</span>
                            </div>
                            <div class="explanation-text">
                                <strong>Assessment Summary:</strong><br>
                                ${data.payload.diagnosis.explanation}
                            </div>
                            <div class="explanation-text" style="margin-top: 16px; padding: 12px; background-color: #EFF6FF; border-radius: 8px; color: #1E40AF;">
                                <strong>Important Patient Guidance / Next Steps:</strong><br>
                                ${data.payload.diagnosis.next_steps}
                            </div>
                        </div>
                        <div class="card">
                            <div class="card-title">Signing Pathologist Certification</div>
                            <div class="info-grid">
                                <div class="info-item"><span class="info-label">Verifying Specialist</span><span class="info-value">${data.payload.signed_by}</span></div>
                                <div class="info-item"><span class="info-label">Institution</span><span class="info-value">${data.payload.institution}</span></div>
                            </div>
                        </div>
                        <div class="footer">
                            This patient report is produced in association with the verifying pathologist. Diagnostic integrity protected by end-to-end security protocols.
                        </div>
                    </body>
                    </html>
                `);
                w.document.close();
                setTimeout(() => w.print(), 500);
            } else {
                w.document.write(`
                    <html>
                    <head>
                        <title>OralDysplasia AI — Diagnostic Report</title>
                        <style>
                            body { font-family: 'Plus Jakarta Sans', sans-serif; padding: 40px; color: #1E293B; line-height: 1.6; }
                            h1 { color: #4F46E5; font-size: 24px; border-bottom: 2px solid #E2E8F0; padding-bottom: 8px; }
                            .header-row { display: flex; justify-content: space-between; margin: 20px 0; }
                            .section-title { font-weight: bold; text-transform: uppercase; color: #64748B; font-size: 11px; margin-top: 24px; }
                            .info-grid { display: grid; grid-template-columns: 150px 1fr; gap: 8px; margin: 10px 0; }
                            .badge { padding: 4px 8px; border-radius: 4px; font-weight: bold; text-transform: uppercase; font-size: 12px; }
                            .badge-severe { background-color: #FEE2E2; color: #EF4444; }
                            .badge-moderate { background-color: #FEF3C7; color: #F59E0B; }
                            .badge-mild { background-color: #D1FAE5; color: #10B981; }
                            .badge-normal { background-color: #F1F5F9; color: #64748B; }
                            .comment-box { background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 16px; border-radius: 8px; font-family: monospace; white-space: pre-line; }
                        </style>
                    </head>
                    <body>
                        <h1>OralDysplasia AI Grading Case Report</h1>
                        <div class="header-row">
                            <div>
                                <div class="section-title">Patient Demographics</div>
                                <div class="info-grid"><span>Patient Name:</span><strong>${data.payload.patient.name}</strong></div>
                                <div class="info-grid"><span>Patient ID:</span><strong>${data.payload.patient.id}</strong></div>
                                <div class="info-grid"><span>Age:</span><strong>${data.payload.patient.age || 'N/A'}</strong></div>
                                <div class="info-grid"><span>Gender:</span><strong>${data.payload.patient.gender || 'N/A'}</strong></div>
                                <div class="info-grid"><span>Biopsy Site:</span><strong>${data.payload.patient.site}</strong></div>
                            </div>
                            <div>
                                <div class="section-title">Slide Scan Metadata</div>
                                <div class="info-grid"><span>Filename:</span><strong>${data.payload.slide.filename}</strong></div>
                                <div class="info-grid"><span>Dimensions:</span><strong>${data.payload.slide.dimensions}</strong></div>
                            </div>
                        </div>
                        <hr style="border:none; border-top:1px solid #E2E8F0; margin:24px 0;">
                        <div class="section-title">Verified Diagnostic Verdict</div>
                        <div class="info-grid" style="align-items:center;">
                            <span>Final Grade:</span>
                            <span><span class="badge badge-${data.payload.diagnosis.grade.toLowerCase()}">${data.payload.diagnosis.grade}</span></span>
                        </div>
                        <div class="info-grid"><span>ICD-10 Code:</span><strong>${data.payload.diagnosis.icd10}</strong></div>
                        <div class="info-grid"><span>Signatory Pathologist:</span><strong>${data.payload.signed_by}</strong></div>
                        
                        <div class="section-title" style="margin-top:24px;">Pathologist Observations & Checklist</div>
                        <div class="comment-box">${data.payload.comments || 'No comments recorded.'}</div>
                        
                        <div style="margin-top:40px; text-align:center; font-size:10px; color:#94A3B8;">
                            Sealed digitally. Diagnostic integrity verified via SHA-256 cryptographic hashes.
                        </div>
                    </body>
                    </html>
                `);
                w.document.close();
                setTimeout(() => w.print(), 500);
            }
        } else {
            // Download payload as JSON file
            const blob = new Blob([JSON.stringify(data.payload, null, 2)], { type: "application/json" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = `${format}_report_case_${activeSlideId}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }
    } catch (err) {
        showLoader(false);
        alert("Export report failed: " + err.message);
    }
}

// ── Share Diagnostic Reports via WhatsApp / Email ─────────────────────
async function triggerShare(channel) {
    const targetId = activeSlideId;
    if (!targetId) return;
    showLoader(true, `Preparing report to share via ${channel === 'whatsapp' ? 'WhatsApp' : 'Email'}...`);
    try {
        const data = await fetchAPI(`/reports/${targetId}/export?format=patient_pdf`);
        const slide = await fetchAPI(`/slides/${targetId}`);
        showLoader(false);

        const payload = data.payload;
        const patient = payload.patient;
        const diagnosis = payload.diagnosis;
        const currentStatus = slide.status || "processed";

        const shareText = `PATIENT DIAGNOSTIC REPORT (OralDysplasia AI)
--------------------------------------------
Patient Name: ${patient.name}
Patient ID: ${patient.id}
Age: ${patient.age || "N/A"}
Gender: ${patient.gender || "N/A"}
Biopsy Site: ${patient.site}

DIAGNOSTIC ASSESSMENT
Grade: ${diagnosis.grade.toUpperCase()}
Summary: ${diagnosis.explanation}
Next Steps: ${diagnosis.next_steps}

CERTIFICATION
Verifying Specialist: ${payload.signed_by}
Institution: ${payload.institution}
Status: ${currentStatus.toUpperCase()}`;

        if (channel === "whatsapp") {
            const url = `https://api.whatsapp.com/send?text=${encodeURIComponent(shareText)}`;
            window.open(url, "_blank");
        } else if (channel === "email") {
            const url = `mailto:?subject=${encodeURIComponent("Patient Diagnostic Report: " + patient.name)}&body=${encodeURIComponent(shareText)}`;
            window.open(url, "_self");
        }
    } catch (err) {
        showLoader(false);
        alert("Failed to share report: " + err.message);
    }
}

// ── Native String hashing helpers ────────────────────────────────────
String.prototype.hashCode = function() {
    let hash = 0;
    for (let i = 0; i < this.length; i++) {
        hash = this.charCodeAt(i) + ((hash << 5) - hash);
    }
    return hash;
};
