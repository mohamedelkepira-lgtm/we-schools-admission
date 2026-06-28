import { Component } from 'react'

export default class ErrorBoundary extends Component {
  constructor(props) {
    super(props)
    this.state = { hasError: false, error: null }
  }

  static getDerivedStateFromError(error) {
    return { hasError: true, error }
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen flex items-center justify-center bg-gray-50 p-4">
          <div className="bg-white rounded-2xl shadow-sm border border-gray-200 p-8 max-w-md text-center">
            <div className="w-14 h-14 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4">
              <svg className="w-7 h-7 text-red-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
            </div>
            <h2 className="text-xl font-bold text-gray-900 mb-2">حدث خطأ غير متوقع</h2>
            <p className="text-gray-500 text-sm mb-4">نأسف للإزعاج. يرجى تحديث الصفحة والمحاولة مرة أخرى.</p>
            <button
              onClick={() => window.location.reload()}
              className="bg-[var(--we-blue)] hover:bg-[var(--we-blue-light)] text-white px-6 py-2.5 rounded-xl font-medium transition-colors"
            >
              تحديث الصفحة
            </button>
            {this.state.error && (
              <details className="mt-4 text-left">
                <summary className="text-xs text-gray-400 cursor-pointer">تفاصيل الخطأ</summary>
                <pre className="text-xs text-red-600 mt-2 bg-red-50 p-2 rounded-lg overflow-auto max-h-32">
                  {this.state.error.message}
                </pre>
              </details>
            )}
          </div>
        </div>
      )
    }
    return this.props.children
  }
}
