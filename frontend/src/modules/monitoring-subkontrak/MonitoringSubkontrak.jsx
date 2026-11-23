import React from 'react';

const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwUG2xVgGE7BDoUJS8xyHwp3pLTQBZqHqpc0gRe5afPbDM8MhMiMvmJPf2c9_N9I00/exec';

export default function MonitoringSubkontrak() {
  return (
    <main style={{ minHeight: '100vh', background: '#f4f7fa', display: 'flex', flexDirection: 'column' }}>
      <header style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 16, padding: '14px 22px', background: '#fff', borderBottom: '1px solid #e2e8ef' }}>
        <div>
          <div style={{ color: '#17365d', fontWeight: 800, fontSize: 17 }}>Monitoring Subkontrak</div>
          <div style={{ color: '#71808d', fontSize: 11, marginTop: 3 }}>REZA SYSTEM · Google Sheets Database</div>
        </div>
        <a href={WEB_APP_URL} target="_blank" rel="noreferrer" style={{ color: '#315c84', fontSize: 12, fontWeight: 750, textDecoration: 'none', whiteSpace: 'nowrap' }}>
          Buka di tab baru ↗
        </a>
      </header>
      <iframe
        title="Monitoring Subkontrak"
        src={WEB_APP_URL}
        allow="clipboard-read; clipboard-write"
        style={{ display: 'block', width: '100%', flex: 1, minHeight: 'calc(100vh - 64px)', border: 0, background: '#fff' }}
      />
    </main>
  );
}
