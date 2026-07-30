import { useState } from 'react'
import axios from 'axios'

const API_BASE = 'http://3.108.190.193:8000'

const OPTIONAL_SECTIONS = [
  { key: 'include_landing_zone',        label: '3.6 Implementation of Landing Zone',         desc: 'OU structure, guardrails, AWS Identity Center setup' },
  { key: 'include_control_tower',       label: '3.7 Configuration of Control Tower Setup',   desc: 'Control Tower automation steps' },
  { key: 'include_landing_zone_arch',   label: '3.8 AWS Landing Zone Architecture',           desc: 'Multi-account architecture details' },
  { key: 'include_paloalto',            label: '3.9 PaloAlto Next Generation Firewall',       desc: 'VM-Series features and capabilities' },
  { key: 'include_mgn_migration',       label: '4.1 Migration of Data Using MGN',             desc: 'AWS Application Migration Service steps' },
  { key: 'include_testing_monitoring',  label: '4.2 Testing and Monitoring',                  desc: 'Post-deployment testing steps and DNS cutover' },
  { key: 'include_monitoring',          label: '5. Monitoring AWS Infrastructure',             desc: 'CloudWatch dashboards, alerts and custom metrics' },
  { key: 'include_dr',                  label: '6. Proposed DR Approach',                     desc: 'DR requirements, AWS DRS, architecture and solution' },
  { key: 'include_post_deployment',     label: '4. Post-Deployment Testing and Acceptance',   desc: 'Client testing window, iteration limits, acceptance policy' },
]

const STATIC_SECTIONS = [
  '1.1 Confidentiality Notice', '1.2 About Operisoft',
  '5.1 AWS TCO Link', '5.2 Service Charges', '5.3 Estimated Timeline',
  '10. Commercial Terms',
]

const DYNAMIC_SECTIONS = [
  '1.3 About Customer', '2.1 Project Objectives', '2.2 Current Landscape',
  '3.2 Key Highlights', '5.4 Cost Assumptions (AI)',
  '6. Customer Obligations and Engagement Terms (AI)',
  '7. Customer Dependencies (AI)',
  '8. Assumptions (AI)',
  '9. Exclusions (AI)',
]

function Toggle({ id, checked, onChange }) {
  return (
    <label className="toggle-switch" htmlFor={id}>
      <input id={id} type="checkbox" checked={checked} onChange={onChange} />
      <span className="slider" />
    </label>
  )
}

export default function App() {
  const [customerName, setCustomerName] = useState('')
  const [companyUrl,   setCompanyUrl]   = useState('')
  const [momText,      setMomText]      = useState('')
  const [projectType,  setProjectType]  = useState('POC')   // 'POC' | 'Production'
  const [docDate,      setDocDate]      = useState(new Date().toISOString().split('T')[0])
  const [submittedBy,  setSubmittedBy]  = useState('')
  const [clientLogo,   setClientLogo]   = useState(null)
  const [logoPreview,  setLogoPreview]  = useState(null)
  const [toggles, setToggles] = useState({
    include_landing_zone:       false,
    include_control_tower:      false,
    include_landing_zone_arch:  false,
    include_paloalto:           false,
    include_mgn_migration:      false,
    include_testing_monitoring: false,
    include_monitoring:         false,
    include_dr:                 false,
    include_post_deployment:    false,
  })
  const [status, setStatus] = useState(null)

  const handleToggle = (key) => setToggles(prev => ({ ...prev, [key]: !prev[key] }))

  const handleLogoChange = (e) => {
    const file = e.target.files[0]
    if (file) {
      setClientLogo(file)
      const reader = new FileReader()
      reader.onload = (ev) => setLogoPreview(ev.target.result)
      reader.readAsDataURL(file)
    }
  }

  const handleLogoRemove = () => {
    setClientLogo(null)
    setLogoPreview(null)
  }

  const handleGenerate = async () => {
    if (!customerName.trim()) { alert('Please enter the customer name.');       return }
    if (!momText.trim())      { alert('Please paste the MOM / meeting notes.'); return }

    setStatus('loading')
    try {
      const formData = new FormData()
      formData.append('customer_name', customerName)
      formData.append('mom_text', momText)
      formData.append('company_url', companyUrl)
      formData.append('project_type', projectType)
      formData.append('doc_date', docDate ? docDate.split('-').reverse().join('.') : '')
      formData.append('submitted_by', submittedBy)
      Object.entries(toggles).forEach(([key, val]) => {
        formData.append(key, val ? 'true' : 'false')
      })
      if (clientLogo) {
        formData.append('client_logo', clientLogo)
      }

      const response = await axios.post(
        `${API_BASE}/generate-sow`,
        formData,
        { responseType: 'blob', headers: { 'Content-Type': 'multipart/form-data' } }
      )
      const url  = window.URL.createObjectURL(new Blob([response.data]))
      const link = document.createElement('a')
      link.href  = url
      link.setAttribute('download', `SOW_${customerName.replace(/\s+/g, '_')}.docx`)
      document.body.appendChild(link)
      link.click()
      link.remove()
      window.URL.revokeObjectURL(url)
      setStatus('success')
    } catch (err) {
      const msg = err.response?.data?.detail || err.message || 'Unknown error'
      setStatus({ error: msg })
    }
  }

  const activeToggles = OPTIONAL_SECTIONS.filter(s => toggles[s.key])

  return (
    <div className="app-shell">

      {/* ── Header ── */}
      <header className="app-header">
        <div className="logo-mark">O</div>
        <h1>Operisoft — SOW Generator</h1>
        <span className="badge">AWS Cloud Services</span>
      </header>

      {/* ── Body ── */}
      <main className="app-body">

        {/* LEFT column */}
        <div>

          {/* Company URL extraction */}
          <div className="card" style={{ marginBottom: '1.25rem' }}>
            <div className="card-title">Customer Information</div>
            <div className="field-group" style={{ marginBottom: '0.75rem' }}>
              <label htmlFor="cname">Customer Name *</label>
              <input
                id="cname"
                type="text"
                placeholder="e.g. Aptech Limited"
                value={customerName}
                onChange={e => setCustomerName(e.target.value)}
              />
            </div>

            {/* Date and Submitted By */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.75rem', marginBottom: '0.75rem' }}>
              <div className="field-group" style={{ marginBottom: 0 }}>
                <label htmlFor="docDate">Document Date *</label>
                <input
                  id="docDate"
                  type="date"
                  value={docDate}
                  onChange={e => setDocDate(e.target.value)}
                />
              </div>
              <div className="field-group" style={{ marginBottom: 0 }}>
                <label htmlFor="submittedBy">Submitted By *</label>
                <input
                  id="submittedBy"
                  type="text"
                  placeholder="Enter name"
                  value={submittedBy}
                  onChange={e => setSubmittedBy(e.target.value)}
                />
              </div>
            </div>

            {/* POC / Production toggle */}
            <div className="field-group" style={{ marginBottom: '0.75rem' }}>
              <label>Engagement Type</label>
              <div className="project-type-toggle">
                <button
                  type="button"
                  className={`ptype-btn${projectType === 'POC' ? ' ptype-active' : ''}`}
                  onClick={() => setProjectType('POC')}
                >
                  POC
                </button>
                <button
                  type="button"
                  className={`ptype-btn${projectType === 'Production' ? ' ptype-active' : ''}`}
                  onClick={() => setProjectType('Production')}
                >
                  Production
                </button>
              </div>
              <p style={{ fontSize: '0.75rem', color: 'var(--muted)', marginTop: '0.3rem', marginBottom: 0 }}>
                Tells the AI whether this is a <strong>Proof of Concept</strong> or a <strong>Production</strong> engagement — affects tone, scope, and disclaimers in the generated SOW.
              </p>
            </div>
            <div className="field-group" style={{ marginBottom: 0 }}>
              <label htmlFor="companyUrl">Company Website URL <span style={{ color: 'var(--muted)', fontWeight: 400 }}>(optional)</span></label>
              <input
                id="companyUrl"
                type="text"
                placeholder="e.g. https://operisoft.com"
                value={companyUrl}
                onChange={e => setCompanyUrl(e.target.value)}
              />
              <p style={{ fontSize: '0.75rem', color: 'var(--muted)', marginTop: '0.3rem', marginBottom: 0 }}>
                Company info will be auto-extracted and used to write the <strong>1.3 About Customer</strong> section.
              </p>
            </div>
          </div>

          {/* Client Logo Upload */}
          <div className="card" style={{ marginBottom: '1.25rem' }}>
            <div className="card-title">Client Logo</div>
            <p style={{ fontSize: '0.8rem', color: 'var(--muted)', marginBottom: '0.75rem' }}>
              Upload the client's logo to include on the SOW cover page.
            </p>
            <div className="logo-upload-area">
              {logoPreview ? (
                <div className="logo-preview-container">
                  <img src={logoPreview} alt="Client logo preview" className="logo-preview-img" />
                  <div className="logo-preview-info">
                    <span className="logo-filename">{clientLogo?.name}</span>
                    <button type="button" className="logo-remove-btn" onClick={handleLogoRemove}>
                      ✕ Remove
                    </button>
                  </div>
                </div>
              ) : (
                <label htmlFor="logoUpload" className="logo-dropzone">
                  <div className="logo-dropzone-icon">🖼️</div>
                  <div className="logo-dropzone-text">Click to upload logo</div>
                  <div className="logo-dropzone-hint">PNG, JPG, SVG — max 5MB</div>
                  <input
                    id="logoUpload"
                    type="file"
                    accept="image/png,image/jpeg,image/svg+xml,image/webp"
                    onChange={handleLogoChange}
                    style={{ display: 'none' }}
                  />
                </label>
              )}
            </div>
            <p style={{ fontSize: '0.72rem', color: 'var(--muted)', marginTop: '0.5rem', marginBottom: 0 }}>
              The logo will appear centered on the cover page, above the customer name — matching the template layout.
            </p>
          </div>

          {/* MOM notes */}
          <div className="card">
            <div className="card-title">Meeting of Minutes (MOM)</div>
            <div className="field-group">
              <label htmlFor="mom">Paste your discovery call notes / MOM here *</label>
              <textarea
                id="mom"
                value={momText}
                onChange={e => setMomText(e.target.value)}
                style={{ minHeight: 320 }}
                placeholder={
                  'Paste meeting notes here. Include:\n' +
                  '• Current infrastructure (servers, OS, databases)\n' +
                  '• Applications running on-premises\n' +
                  '• Migration goals and timelines\n' +
                  '• DR requirements (RTO / RPO)\n' +
                  '• Specific AWS services requested\n' +
                  '• Compliance or security requirements'
                }
              />
            </div>
            <p style={{ fontSize: '0.78rem', color: 'var(--muted)' }}>
              AWS Bedrock Claude Haiku uses these notes to generate the dynamic SOW sections.
              Credentials are loaded from <code style={{ background: '#f1f5f9', padding: '1px 5px', borderRadius: 4 }}>backend/.env</code>.
            </p>
          </div>

        </div>

        {/* RIGHT column */}
        <div>

          {/* Optional toggles */}
          <div className="card" style={{ marginBottom: '1.25rem' }}>
            <div className="card-title">Optional Sections</div>
            <p style={{ fontSize: '0.8rem', color: 'var(--muted)', marginBottom: '1rem' }}>
              Toggle sections that apply to this engagement.
            </p>
            {OPTIONAL_SECTIONS.map(sec => (
              <div key={sec.key} className={`toggle-row${toggles[sec.key] ? ' active' : ''}`}>
                <div className="toggle-info">
                  <div className="toggle-label">{sec.label}</div>
                  <div className="toggle-desc">{sec.desc}</div>
                </div>
                <Toggle
                  id={sec.key}
                  checked={toggles[sec.key]}
                  onChange={() => handleToggle(sec.key)}
                />
              </div>
            ))}
          </div>

          {/* Generate */}
          <div className="card">
            <div className="card-title">Generate SOW</div>

            <div style={{ fontSize: '0.8rem', color: 'var(--muted)', marginBottom: '0.75rem' }}>
              Sections included:&nbsp;
              <strong style={{ color: 'var(--navy)' }}>
                {STATIC_SECTIONS.length + DYNAMIC_SECTIONS.length + activeToggles.length}
              </strong>
              &nbsp;({activeToggles.length} optional active)
            </div>

            <button
              className="btn-generate"
              onClick={handleGenerate}
              disabled={status === 'loading'}
            >
              {status === 'loading'
                ? <><span className="spinner" /> Generating SOW…</>
                : <><span>⬇</span> Generate &amp; Download SOW</>}
            </button>

            {status === 'loading' && (
              <div className="status-box loading">
                <span className="spinner" />
                Calling Bedrock for dynamic content — takes 20–40 seconds…
              </div>
            )}
            {status === 'success' && (
              <div className="status-box success">✓ SOW generated and downloaded!</div>
            )}
            {status?.error && (
              <div className="status-box error">✗ {status.error}</div>
            )}

            <div className="divider" />

            {/* Sections legend */}
            <div className="card-title" style={{ fontSize: '0.75rem' }}>Sections Overview</div>
            <div className="legend">
              <span className="legend-item">
                <span className="legend-dot" style={{ background: '#6366f1' }} /> Static
              </span>
              <span className="legend-item">
                <span className="legend-dot" style={{ background: 'var(--orange)' }} /> AI Generated
              </span>
              <span className="legend-item">
                <span className="legend-dot" style={{ background: '#16a34a' }} /> Optional (on)
              </span>
            </div>
            <div>
              {STATIC_SECTIONS.map(s  => <span key={s}     className="section-badge static">{s}</span>)}
              {DYNAMIC_SECTIONS.map(s => <span key={s}     className="section-badge dynamic">✦ {s}</span>)}
              {activeToggles.map(s    => <span key={s.key} className="section-badge optional">✓ {s.label}</span>)}
            </div>
          </div>

        </div>
      </main>
    </div>
  )
}
