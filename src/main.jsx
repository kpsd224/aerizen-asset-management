import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.jsx';
import './styles.css';

class AppErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }

  componentDidCatch(error) {
    this.setState({ error });
  }

  resetLocalData = () => {
    localStorage.removeItem('aerizen.v3.state');
    localStorage.removeItem('aerizen.v3.settings');
    localStorage.removeItem('aerizen.v3.syncQueue');
    window.location.reload();
  };

  render() {
    if (this.state.error) {
      return (
        <div style={{ minHeight: '100vh', display: 'grid', placeItems: 'center', background: '#0f172a', color: '#e5e7eb', fontFamily: 'Arial, sans-serif', padding: 24 }}>
          <div style={{ maxWidth: 760, background: '#111827', border: '1px solid #334155', borderRadius: 24, padding: 28 }}>
            <p style={{ color: '#22d3ee', fontWeight: 800, letterSpacing: '.08em', textTransform: 'uppercase' }}>Aerizen Recovery Mode</p>
            <h1 style={{ margin: '8px 0 12px', fontSize: 30 }}>Aplikasi gagal dirender</h1>
            <p style={{ color: '#cbd5e1', lineHeight: 1.7 }}>
              Ini biasanya terjadi karena data lokal dari versi lama tidak cocok dengan struktur versi baru. Data Supabase tetap aman. Coba tutup aplikasi dan buka ulang. Jika masih gagal, reset data lokal demo di perangkat ini.
            </p>
            <pre style={{ whiteSpace: 'pre-wrap', background: '#020617', color: '#fecaca', padding: 16, borderRadius: 16 }}>{String(this.state.error?.message || this.state.error)}</pre>
            <button onClick={this.resetLocalData} style={{ border: 0, borderRadius: 14, padding: '12px 18px', fontWeight: 800, background: '#dc2626', color: '#fff', cursor: 'pointer' }}>Reset data lokal perangkat ini</button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

window.addEventListener('error', (event) => {
  console.error('Aerizen runtime error:', event.error || event.message);
});
window.addEventListener('unhandledrejection', (event) => {
  console.error('Aerizen unhandled rejection:', event.reason);
});

createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <AppErrorBoundary>
      <App />
    </AppErrorBoundary>
  </React.StrictMode>
);
