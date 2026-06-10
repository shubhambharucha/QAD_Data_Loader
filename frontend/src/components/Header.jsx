import "./Header.css";

export default function Header() {
  return (
    <header className="header">
      <div className="header-inner">
        <div className="header-logo-wrap">
          <img
            src="/logo.png"
            alt="QAD logo"
            className="header-logo"
            onError={e => { e.target.style.display = "none"; }}
          />
        </div>
        <div className="header-title-wrap">
          <h1 className="header-title">QAD Data Loader</h1>
          <span className="header-sub">Enterprise Data Operations</span>
        </div>
      </div>
      <div className="header-divider" />
    </header>
  );
}