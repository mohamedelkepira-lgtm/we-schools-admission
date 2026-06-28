import { lazy, Suspense } from 'react'
import { Routes, Route } from 'react-router-dom'
import ErrorBoundary from './components/ErrorBoundary'
import LoadingFallback from './components/LoadingFallback'
import Layout from './components/Layout'

const Home = lazy(() => import('./pages/Home'))
const Documents = lazy(() => import('./pages/Documents'))
const Register = lazy(() => import('./pages/Register'))
const Success = lazy(() => import('./pages/Success'))

export default function App() {
  return (
    <ErrorBoundary>
      <Layout>
        <Suspense fallback={<LoadingFallback />}>
          <Routes>
            <Route path="/" element={<Home />} />
            <Route path="/documents" element={<Documents />} />
            <Route path="/register" element={<Register />} />
            <Route path="/success" element={<Success />} />
          </Routes>
        </Suspense>
      </Layout>
    </ErrorBoundary>
  )
}
