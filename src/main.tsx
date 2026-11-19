import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.tsx'

const canUseServiceWorker =
  'serviceWorker' in navigator && ['http:', 'https:'].includes(window.location.protocol)

if (canUseServiceWorker) {
  window.addEventListener('load', () => {
    const serviceWorkerUrl = new URL(`${import.meta.env.BASE_URL}sw.js`, window.location.href)

    navigator.serviceWorker.register(serviceWorkerUrl).catch(() => {
      // Keep the app usable even if local web-app registration is unavailable.
    })
  })
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
