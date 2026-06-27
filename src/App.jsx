import { Routes, Route } from 'react-router-dom'
import Layout from './components/Layout'
import Home from './pages/Home'
import Documents from './pages/Documents'
import Register from './pages/Register'
import Success from './pages/Success'

export default function App() {
  return (
    <Layout>
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/documents" element={<Documents />} />
        <Route path="/register" element={<Register />} />
        <Route path="/success" element={<Success />} />
      </Routes>
    </Layout>
  )
}
