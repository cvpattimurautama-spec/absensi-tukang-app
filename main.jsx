import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App.jsx';
import './index.css';

function ErrorScreen({ title, detail }) {
  return (
    <div style={{ padding: 24, fontFamily: 'monospace', background: '#2a1414', color: '#ffb4a8', minHeight: '100vh' }}>
      <h2 style={{ color: '#ff6b5b', marginTop: 0 }}>⚠ {title}</h2>
      <pre style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', background: '#1a0a0a', padding: 12, borderRadius: 6, fontSize: 13 }}>{detail}</pre>
      <p style={{ fontSize: 13, color: '#e8a08f' }}>
        Screenshot halaman ini dan kirim ke Claude untuk dibantu diagnosa.
      </p>
    </div>
  );
}

class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error) {
    return { error };
  }
  componentDidCatch(error, info) {
    console.error('Render error:', error, info);
  }
  render() {
    if (this.state.error) {
      return <ErrorScreen title="Terjadi error saat menampilkan aplikasi" detail={String(this.state.error?.stack || this.state.error)} />;
    }
    return this.props.children;
  }
}

// Tangkap error yang terjadi SEBELUM React sempat mengambil alih (mis. saat memuat modul)
window.addEventListener('error', (e) => {
  const root = document.getElementById('root');
  if (root && !root.hasChildNodes()) {
    ReactDOM.createRoot(root).render(<ErrorScreen title="Error saat memuat aplikasi" detail={String(e.error?.stack || e.message || e)} />);
  }
});
window.addEventListener('unhandledrejection', (e) => {
  const root = document.getElementById('root');
  if (root && !root.hasChildNodes()) {
    ReactDOM.createRoot(root).render(<ErrorScreen title="Error saat memuat aplikasi (async)" detail={String(e.reason?.stack || e.reason || e)} />);
  }
});

try {
  ReactDOM.createRoot(document.getElementById('root')).render(
    <React.StrictMode>
      <ErrorBoundary>
        <App />
      </ErrorBoundary>
    </React.StrictMode>
  );
} catch (err) {
  ReactDOM.createRoot(document.getElementById('root')).render(
    <ErrorScreen title="Error saat memulai aplikasi" detail={String(err?.stack || err)} />
  );
}
