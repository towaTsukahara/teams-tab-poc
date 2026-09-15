import { useEffect, useState } from 'react'
import { app } from '@microsoft/teams-js'
import logo from './assets/logo.png'
import './App.css'

type ViewKey = 'dashboard' | 'projects' | 'tasks' | 'reports' | 'settings'

const NAV_ITEMS: { key: ViewKey; label: string }[] = [
  { key: 'dashboard', label: 'ダッシュボード' },
  { key: 'projects', label: 'プロジェクト' },
  { key: 'tasks', label: 'タスク' },
  { key: 'reports', label: 'レポート' },
  { key: 'settings', label: '設定' },
]

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
  { name: 'セキュリティ監査対応', status: '進行中', owner: '伊藤', due: '2026-10-05' },
  { name: 'コールセンターAI導入', status: '遅延', owner: '渡辺', due: '2026-09-25' },
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

const TASKS = [
  { title: '要件定義書のレビュー', owner: '佐藤', done: true },
  { title: 'API設計書の作成', owner: '鈴木', done: false },
  { title: 'テスト計画の策定', owner: '高橋', done: false },
  { title: 'ステークホルダー向け報告資料', owner: '田中', done: true },
  { title: 'セキュリティ診断の日程調整', owner: '伊藤', done: false },
]

function App() {
  const [status, setStatus] = useState<'loading' | 'in-teams' | 'browser'>('loading')
  const [theme, setTheme] = useState<string>('default')
  const [userName, setUserName] = useState<string | null>(null)
  const [activeView, setActiveView] = useState<ViewKey>('dashboard')

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
          {NAV_ITEMS.map((item) => (
            <div
              key={item.key}
              className={`nav-item ${activeView === item.key ? 'active' : ''}`}
              role="button"
              tabIndex={0}
              onClick={() => setActiveView(item.key)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') setActiveView(item.key)
              }}
            >
              {item.label}
            </div>
          ))}
        </nav>

        <main className="main-content">
          {activeView === 'dashboard' && <DashboardView />}
          {activeView === 'projects' && <ProjectsView />}
          {activeView === 'tasks' && <TasksView />}
          {activeView === 'reports' && <ReportsView />}
          {activeView === 'settings' && (
            <SettingsView theme={theme} userName={userName} status={status} />
          )}
        </main>
      </div>
    </div>
  )
}

function DashboardView() {
  return (
    <>
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
          <ProjectTable rows={PROJECTS.slice(0, 4)} />
        </div>

        <div className="panel">
          <h2>月次アクティビティ</h2>
          <BarChart />
        </div>
      </section>
    </>
  )
}

function ProjectsView() {
  return (
    <>
      <h1>プロジェクト</h1>
      <p className="subtitle">全{PROJECTS.length}件のプロジェクトを表示しています。</p>
      <section className="panel">
        <ProjectTable rows={PROJECTS} />
      </section>
    </>
  )
}

function ProjectTable({ rows }: { rows: typeof PROJECTS }) {
  return (
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
        {rows.map((p) => (
          <tr key={p.name}>
            <td>{p.name}</td>
            <td>
              <span className={`status-badge ${STATUS_CLASS[p.status]}`}>{p.status}</span>
            </td>
            <td>{p.owner}</td>
            <td>{p.due}</td>
          </tr>
        ))}
      </tbody>
    </table>
  )
}

function TasksView() {
  const doneCount = TASKS.filter((t) => t.done).length
  return (
    <>
      <h1>タスク</h1>
      <p className="subtitle">
        {TASKS.length}件中 {doneCount}件完了
      </p>
      <section className="panel">
        <ul className="task-list">
          {TASKS.map((t) => (
            <li key={t.title} className={`task-item ${t.done ? 'done' : ''}`}>
              <span className="task-check">{t.done ? '✓' : ''}</span>
              <span className="task-title">{t.title}</span>
              <span className="task-owner">{t.owner}</span>
            </li>
          ))}
        </ul>
      </section>
    </>
  )
}

function ReportsView() {
  return (
    <>
      <h1>レポート</h1>
      <p className="subtitle">月次のアクティビティ推移（モックデータ）</p>
      <section className="content-grid">
        <div className="panel">
          <h2>月次アクティビティ</h2>
          <BarChart />
        </div>
        <div className="panel">
          <h2>サマリー</h2>
          <ul className="summary-list">
            <li>平均完了率: 68%</li>
            <li>最も活発だった月: 9月</li>
            <li>遅延プロジェクト: 2件</li>
          </ul>
          <button type="button" className="mock-button" disabled>
            レポートをエクスポート（検証用のため無効）
          </button>
        </div>
      </section>
    </>
  )
}

function SettingsView({
  theme,
  userName,
  status,
}: {
  theme: string
  userName: string | null
  status: 'loading' | 'in-teams' | 'browser'
}) {
  return (
    <>
      <h1>設定</h1>
      <p className="subtitle">検証用アプリのため、変更は保存されません。</p>
      <section className="panel settings-panel">
        <div className="settings-row">
          <span>表示名</span>
          <span className="settings-value">{userName ?? '(Teams外のため取得なし)'}</span>
        </div>
        <div className="settings-row">
          <span>読み込み状態</span>
          <span className="settings-value">
            {status === 'in-teams' ? 'Teams内で読み込み中' : 'ブラウザ単体で表示中'}
          </span>
        </div>
        <div className="settings-row">
          <span>現在のテーマ</span>
          <span className="settings-value">{theme}</span>
        </div>
        <div className="settings-row">
          <span>通知</span>
          <span className="toggle on" />
        </div>
        <div className="settings-row">
          <span>週次レポートの自動送信</span>
          <span className="toggle" />
        </div>
      </section>
    </>
  )
}

function BarChart() {
  return (
    <div className="bar-chart">
      {ACTIVITY.map((a) => (
        <div key={a.label} className="bar-col">
          <div className="bar" style={{ height: `${a.value}%` }} />
          <div className="bar-label">{a.label}</div>
        </div>
      ))}
    </div>
  )
}

export default App
