import { useEffect, useState } from 'react'
import { app } from '@microsoft/teams-js'
import logo from './assets/logo.png'
import './App.css'

const NAV_ITEMS = ['ダッシュボード', 'プロジェクト', 'タスク', 'レポート', '設定']

const KPI_CARDS = [
  { label: '進行中プロジェクト', value: '12', delta: '+2', deltaType: 'up' as const },
  { label: '今月のタスク完了', value: '87', delta: '+15', deltaType: 'up' as const },
  { label: '未対応アラート', value: '3', delta: '-1', deltaType: 'down' as const },
  { label: 'チーム稼働率', value: '92%', delta: '+4%', deltaType: 'up' as const },
]

const PROJECTS = [
  { name: 'クラウド基盤移行', status: '進行中', owner: '佐藤', due: '2026-09-30' },
  { name: 'AI需要予測PoC', status: '進行中', owner: '鈴木', due: '2026-10-15' },
  { name: 'ICT資産棚卸し', status: '遅延', owner: '高橋', due: '2026-09-20' },
  { name: '社内ポータル刷新', status: '完了', owner: '田中', due: '2026-09-10' },
]

const STATUS_CLASS: Record<string, string> = {
  進行中: 'status-progress',
  完了: 'status-done',
  遅延: 'status-delayed',
}

const ACTIVITY = [
  { label: '4月', value: 40 },
  { label: '5月', value: 55 },
  { label: '6月', value: 48 },
  { label: '7月', value: 70 },
  { label: '8月', value: 62 },
  { label: '9月', value: 87 },
]

function App() {
  const [status, setStatus] = useState<'loading' | 'in-teams' | 'browser'>('loading')
  const [theme, setTheme] = useState<string>('default')
  const [userName, setUserName] = useState<string | null>(null)

  useEffect(() => {
    app
      .initialize()
      .then(async () => {
        const ctx = await app.getContext()
        setTheme(ctx.app.theme)
        setUserName(ctx.user?.displayName ?? null)
        setStatus('in-teams')
        app.registerOnThemeChangeHandler((newTheme) => setTheme(newTheme))
      })
      .catch(() => setStatus('browser'))
  }, [])

  return (
    <div className={`app-shell theme-${theme}`}>
      <header className="topbar">
        <div className="brand">
          <img src={logo} alt="AI CLOUD ICT" className="brand-logo" />
          <span className="brand-name">AI CLOUD ICT ダッシュボード</span>
        </div>
        <div className="topbar-user">
          {status === 'loading' && '読み込み中...'}
          {status === 'browser' && 'ゲスト（ブラウザ表示）'}
          {status === 'in-teams' && (userName ?? 'Teamsユーザー')}
        </div>
      </header>

      <div className="body-layout">
        <nav className="sidebar">
          {NAV_ITEMS.map((item, i) => (
            <div key={item} className={`nav-item ${i === 0 ? 'active' : ''}`}>
              {item}
            </div>
          ))}
        </nav>

        <main className="main-content">
          <h1>ダッシュボード</h1>
          <p className="subtitle">検証用のモックデータを表示しています。</p>

          <section className="kpi-row">
            {KPI_CARDS.map((card) => (
              <div key={card.label} className="kpi-card">
                <div className="kpi-label">{card.label}</div>
                <div className="kpi-value">{card.value}</div>
                <div className={`kpi-delta ${card.deltaType}`}>{card.delta}</div>
              </div>
            ))}
          </section>

          <section className="content-grid">
            <div className="panel">
              <h2>最近のプロジェクト</h2>
              <table className="project-table">
                <thead>
                  <tr>
                    <th>プロジェクト名</th>
                    <th>ステータス</th>
                    <th>担当</th>
                    <th>期限</th>
                  </tr>
                </thead>
                <tbody>
                  {PROJECTS.map((p) => (
                    <tr key={p.name}>
                      <td>{p.name}</td>
                      <td>
                        <span className={`status-badge ${STATUS_CLASS[p.status]}`}>
                          {p.status}
                        </span>
                      </td>
                      <td>{p.owner}</td>
                      <td>{p.due}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="panel">
              <h2>月次アクティビティ</h2>
              <div className="bar-chart">
                {ACTIVITY.map((a) => (
                  <div key={a.label} className="bar-col">
                    <div className="bar" style={{ height: `${a.value}%` }} />
                    <div className="bar-label">{a.label}</div>
                  </div>
                ))}
              </div>
            </div>
          </section>
        </main>
      </div>
    </div>
  )
}

export default App
