import { Link, useLocation } from 'react-router-dom'

const links = [
  { to: '/', label: 'الرئيسية' },
  { to: '/documents', label: 'الأوراق المطلوبة' },
  { to: '/register', label: 'التسجيل' },
]

export default function Navbar() {
  const { pathname } = useLocation()

  return (
    <nav className="bg-white shadow-sm border-b border-gray-200 sticky top-0 z-50">
      <div className="max-w-6xl mx-auto px-4 h-16 flex items-center justify-between">
        <Link to="/" className="text-xl font-bold text-[var(--we-blue)]">
          WE <span className="text-[var(--we-gold)]">Schools</span>
        </Link>
        <div className="flex gap-1">
          {links.map((link) => {
            const isActive = pathname === link.to
            return (
              <Link
                key={link.to}
                to={link.to}
                className={`relative px-4 py-2 rounded-lg text-sm font-medium transition-colors duration-200 ${
                  isActive
                    ? 'text-[var(--we-blue)]'
                    : 'text-gray-600 hover:text-[var(--we-blue)] hover:bg-gray-50'
                }`}
              >
                {link.label}
                {isActive && (
                  <span className="absolute bottom-0 left-1/2 -translate-x-1/2 w-6 h-0.5 bg-[var(--we-gold)] rounded-full" />
                )}
              </Link>
            )
          })}
        </div>
      </div>
    </nav>
  )
}
