document.addEventListener('DOMContentLoaded', function () {
    // Initialize tooltips
    const tooltips = document.querySelectorAll('[data-bs-toggle="tooltip"]');
    tooltips.forEach(el => new bootstrap.Tooltip(el));

    // --- Form Submission ---
    const form = document.getElementById('underwritingForm');
    if (!form) return;

    form.addEventListener('submit', function (e) {
        e.preventDefault();

        // Validate required fields
        const requiredFields = form.querySelectorAll('[required]');
        let valid = true;
        requiredFields.forEach(field => {
            if (!field.value.trim()) {
                field.classList.add('is-invalid');
                valid = false;
            } else {
                field.classList.remove('is-invalid');
            }
        });
        if (!valid) return;

        const submitBtn = document.getElementById('submitBtn');
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Running analysis...';

        const fd = new FormData(form);

        fetch('/api/analyze', {
            method: 'POST',
            body: fd
        })
        .then(r => r.json())
        .then(data => {
            if (data.job_id) {
                window.location.href = '/processing/' + data.job_id;
            } else if (data.error) {
                alert('Error: ' + data.error);
                submitBtn.disabled = false;
                submitBtn.innerHTML = '<i class="bi bi-play-fill"></i> Run Underwriting Analysis';
            }
        })
        .catch(err => {
            alert('Network error: ' + err.message);
            submitBtn.disabled = false;
            submitBtn.innerHTML = '<i class="bi bi-play-fill"></i> Run Underwriting Analysis';
        });
    });

    // Clear validation on input
    form.querySelectorAll('.form-control, .form-select').forEach(el => {
        el.addEventListener('input', function () {
            this.classList.remove('is-invalid');
        });
    });

    // --- Load Demo Values ---
    const demoBtn = document.getElementById('loadDemoBtn');
    if (demoBtn) {
        demoBtn.addEventListener('click', function () {
            const demo = {
                property_type: 'Multifamily - Class B',
                year_built: '1995',
                address: '4500 Duval Street, Austin, TX 78751',
                purchase_price: '12500000',
                current_noi: '650000',
                total_units: '50',
                total_sf: '45000',
                in_place_rent: '1300',
                market_rent: '1450',
                occupancy: '92',
                deferred_maintenance: '200000',
                planned_capex: '500000',
                capex_description: 'Unit upgrades',
                hold_period_years: '7',
                // Tax defaults
                tax_rate: '25',
                land_value_pct: '20',
            };

            Object.entries(demo).forEach(([name, value]) => {
                const el = form.querySelector(`[name="${name}"]`);
                if (!el) return;
                el.value = value;
                el.classList.remove('is-invalid');
            });

            // Enable ML features for demo
            const mlCheck = form.querySelector('#mlValuation');
            const rentCheck = form.querySelector('#rentPrediction');
            if (mlCheck) mlCheck.checked = true;
            if (rentCheck) rentCheck.checked = true;

            // Brief visual feedback
            demoBtn.innerHTML = '<i class="bi bi-check-lg"></i> Loaded!';
            demoBtn.classList.replace('btn-outline-warning', 'btn-warning');
            setTimeout(() => {
                demoBtn.innerHTML = '<i class="bi bi-lightning-fill"></i> Load Demo Values';
                demoBtn.classList.replace('btn-warning', 'btn-outline-warning');
            }, 1500);
        });
    }
});
