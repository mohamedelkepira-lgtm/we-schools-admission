export default function ProgressBar({ percentage }) {
  const clamped = Math.min(100, Math.max(0, percentage))

  return (
    <div className="w-full">
      <div className="flex justify-between text-sm text-gray-600 mb-1.5">
        <span>اكتمال النموذج</span>
        <span className="font-medium text-[var(--we-blue)]">{Math.round(clamped)}%</span>
      </div>
      <div className="w-full h-3 bg-gray-200 rounded-full overflow-hidden">
        <div
          className="h-full bg-gradient-to-l from-[var(--we-blue)] to-[var(--we-blue-light)] rounded-full transition-all duration-700 ease-out"
          style={{ width: `${clamped}%` }}
        />
      </div>
    </div>
  )
}
