import { useEffect, useState } from 'react'
import { app } from '@microsoft/teams-js'
import './App.css'

function App() {
  const [status, setStatus] = useState<'loading' | 'in-teams' | 'browser'>('loading')
  const [theme, setTheme] = useState<string>('default')
  const [context, setContext] = useState<app.Context | null>(null)

  useEffect(() => {
    app
      .initialize()
      .then(async () => {
        const ctx = await app.getContext()
        setContext(ctx)
        setTheme(ctx.app.theme)
        setStatus('in-teams')

        app.registerOnThemeChangeHandler((newTheme) => {
          setTheme(newTheme)
        })
      })
      .catch(() => {
        setStatus('browser')
      })
  }, [])

  return (
    <div className={`page theme-${theme}`}>
      <h1>Teams タブ 検証用アプリ</h1>

      {status === 'loading' && <p>読み込み中...</p>}

      {status === 'browser' && (
        <p>
          Teams外（ブラウザ）から開いています。Teamsにサイドロードすると
          <code>app.initialize()</code> が成功し、コンテキスト情報が表示されます。
        </p>
      )}

      {status === 'in-teams' && context && (
        <div className="context-box">
          <p>Teams内での読み込みに成功しました。</p>
          <ul>
            <li>テーマ: {theme}</li>
            <li>ユーザー表示名: {context.user?.displayName ?? '(取得なし)'}</li>
            <li>テナントID: {context.user?.tenant?.id ?? '(取得なし)'}</li>
          </ul>
        </div>
      )}
    </div>
  )
}

export default App
