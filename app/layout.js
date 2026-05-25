import './globals.css'

export const metadata = {
  title: 'Трекер поставок',
  description: 'Учёт поставок из Китая',
}

export default function RootLayout({ children }) {
  return (
    <html lang="ru">
      <body>{children}</body>
    </html>
  )
}
