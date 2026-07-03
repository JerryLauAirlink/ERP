# Re-apply AIRLINK client layer after copying DEMO -> AIRLINK
$path = Join-Path $PSScriptRoot "..\AIRLINK\index.html"
$html = Get-Content $path -Raw -Encoding UTF8

$html = $html -replace '<title>ERP System \(DEV\)</title>', '<title>AIRLINK ERP</title>'
$html = $html -replace 'const ERP_DEV_MODE = true;', @'
const AIRLINK_BRAND = {
        logo: "/AIRLINK/logo.png",
        loginScreen: "/AIRLINK/login-screen.png",
        name: "AIRLINK"
      };
'@

$html = $html -replace 'appTitle: "ERP System \(DEV\)"', 'appTitle: "AIRLINK ERP"'
$html = $html -replace 'appTitle: "ERP 系統（開發）"', 'appTitle: "AIRLINK ERP"'

$html = $html -replace 'localStorage\.getItem\("erp_company_name"\) \|\| ""', 'localStorage.getItem("erp_company_name") || AIRLINK_BRAND.name'
$html = $html -replace 'localStorage\.getItem\("erp_company_logo"\) \|\| ""', 'localStorage.getItem("erp_company_logo") || AIRLINK_BRAND.logo'

# Inject login/auth wrapper before function App()
$loginBlock = @'

      function LoginPage({ onLogin }) {
        const lang = localStorage.getItem("erp_lang") || "zh_TW";
        const t = (key) => (I18N[i18nLang(lang)] && I18N[i18nLang(lang)][key]) || I18N.en[key] || key;
        const [username, setUsername] = useState(() => localStorage.getItem("erp_saved_username") || "");
        const [password, setPassword] = useState("");
        const [remember, setRemember] = useState(() => localStorage.getItem("erp_remember_login") === "1");
        const [showPw, setShowPw] = useState(false);
        const [error, setError] = useState("");
        const [submitting, setSubmitting] = useState(false);
        function handleSubmit(e) {
          e.preventDefault();
          setError("");
          const id = String(username || "").trim();
          const pw = String(password || "");
          if (!id || !pw) { setError(t("loginFail")); return; }
          setSubmitting(true);
          const users = loadJson("erp_users", initialUsers);
          const found = users.find((u) => {
            const email = String(u.email || "").trim().toLowerCase();
            const name = String(u.name || "").trim().toLowerCase();
            const key = id.toLowerCase();
            return (email === key || name === key) && String(u.password || "") === pw;
          });
          if (!found) { setError(t("loginFail")); setSubmitting(false); return; }
          if (!found.is_active) { setError(t("loginInactive")); setSubmitting(false); return; }
          if (remember) localStorage.setItem("erp_saved_username", id);
          else localStorage.removeItem("erp_saved_username");
          persistAuthSession(found.id, remember);
          onLogin(found.id);
        }
        return (
          <div className="erp-login-shell">
            <div className="erp-login-frame">
              <img src={AIRLINK_BRAND.loginScreen} alt="AIRLINK ERP" />
              <form onSubmit={handleSubmit} className="erp-login-overlay" aria-label="Login">
                <div className="erp-login-field">
                  <input type="text" autoComplete="username" placeholder={t("loginUsername")} value={username} onChange={(e) => setUsername(e.target.value)} />
                </div>
                <div className="erp-login-field">
                  <input type={showPw ? "text" : "password"} autoComplete="current-password" placeholder={t("loginPassword")} value={password} onChange={(e) => setPassword(e.target.value)} />
                  <button type="button" onClick={() => setShowPw((v) => !v)} className="text-slate-400 hover:text-slate-600 p-1 shrink-0" tabIndex={-1} aria-label="Toggle password">
                    {showPw ? (
                      <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.75" d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858 3.05a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242M9.88 9.88l-3.29-3.29m7.532 7.532l3.29 3.29M3 3l3.59 3.59m0 0A9.953 9.953 0 0112 5c4.478 0 8.268 2.943 9.543 7a10.025 10.025 0 01-4.132 5.411m0 0L21 21" /></svg>
                    ) : (
                      <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.75" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" /><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.75" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" /></svg>
                    )}
                  </button>
                </div>
                <label className="erp-login-remember cursor-pointer select-none">
                  <input type="checkbox" className="rounded border-slate-300 text-blue-600" checked={remember} onChange={(e) => setRemember(e.target.checked)} />
                  {t("loginRemember")}
                </label>
                {error && <p className="erp-login-error">{error}</p>}
                <button type="submit" className="erp-login-btn" disabled={submitting} aria-label={t("loginBtn")} title={t("loginBtn")} />
              </form>
            </div>
          </div>
        );
      }

      function resolveAuthUserId() {
        const raw = sessionStorage.getItem("erp_session_user_id")
          || (localStorage.getItem("erp_remember_login") === "1" ? localStorage.getItem("erp_session_user_id") : null);
        if (!raw) return null;
        const users = loadJson("erp_users", initialUsers);
        const u = users.find((x) => x.id === Number(raw) && x.is_active);
        return u ? u.id : null;
      }

      function persistAuthSession(userId, remember) {
        sessionStorage.setItem("erp_session_user_id", String(userId));
        localStorage.setItem("erp_current_user_id", String(userId));
        if (remember) {
          localStorage.setItem("erp_remember_login", "1");
          localStorage.setItem("erp_session_user_id", String(userId));
        } else {
          localStorage.removeItem("erp_remember_login");
          localStorage.removeItem("erp_session_user_id");
        }
      }

      function clearAuthSession() {
        sessionStorage.removeItem("erp_session_user_id");
        localStorage.removeItem("erp_session_user_id");
        localStorage.removeItem("erp_remember_login");
      }

      function ErpRoot() {
        const [authUserId, setAuthUserId] = useState(() => resolveAuthUserId());
        if (!authUserId) {
          return <LoginPage onLogin={(id) => setAuthUserId(id)} />;
        }
        return (
          <App
            authUserId={authUserId}
            onLogout={() => {
              clearAuthSession();
              setAuthUserId(null);
            }}
          />
        );
      }

'@

if ($html -notmatch 'function LoginPage') {
  $html = $html -replace '(\r?\n      function App\(\) \{)', ($loginBlock + '$1')
}

$html = $html -replace 'function App\(\) \{', 'function App({ authUserId, onLogout }) {'
$html = $html -replace 'useState\(\(\) => Number\(localStorage\.getItem\("erp_current_user_id"\) \|\| 1\)\)', 'useState(() => authUserId || Number(localStorage.getItem("erp_current_user_id") || 1))'

$oldUserUi = @'
                  <div className="flex items-center gap-1.5 h-9">
                    <span className="text-[10px] text-slate-400 whitespace-nowrap">{t("currentUser")}</span>
                    <select value={currentUserId} onChange={(e) => setCurrentUserId(Number(e.target.value))} className="h-9 min-w-[7rem] rounded-lg border border-slate-300 bg-white text-xs px-2" title={t("switchUser")}>
                      {users.filter((u) => u.is_active).map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}
                    </select>
                  </div>
'@

$newUserUi = @'
                  <div className="flex items-center gap-1.5 h-9 px-2.5 rounded-lg border border-slate-200 bg-white">
                    <span className="text-[10px] text-slate-400 whitespace-nowrap">{t("currentUser")}</span>
                    <span className="text-xs font-semibold text-slate-700 whitespace-nowrap max-w-[8rem] truncate">{getCurrentUser() ? getCurrentUser().name : "—"}</span>
                    <span className="text-slate-300">|</span>
                    <button type="button" onClick={onLogout} className="text-xs text-slate-600 hover:text-red-600 whitespace-nowrap">{t("logout")}</button>
                  </div>
'@

$html = $html.Replace($oldUserUi.Trim(), $newUserUi.Trim())
$html = $html -replace 'ReactDOM\.createRoot\(document\.getElementById\("app"\)\)\.render\(<App />\);', 'ReactDOM.createRoot(document.getElementById("app")).render(<ErpRoot />);'

Set-Content $path $html -Encoding UTF8 -NoNewline
Write-Host "AIRLINK client layer applied."
