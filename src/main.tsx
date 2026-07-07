import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.tsx'
import { bootstrapM365Auth } from './services/connectors'

// Finish Microsoft sign-in in popup before rendering the full app
bootstrapM365Auth().finally(() => {
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <App />
    </StrictMode>,
  )
})