import './globals.css'

export const metadata = {
  title: 'Трекер поставок',
  description: 'Учёт поставок из Китая',
  manifest: '/manifest.json',
  themeColor: '#0F4C75',
  appleWebApp: {
    capable: true,
    statusBarStyle: 'black-translucent',
    title: 'Трекер',
  },
  icons: {
    apple: '/icon.svg',
  },
}

export default function RootLayout({ children }) {
  return (
    <html lang="ru">
      <head>
        <link rel="apple-touch-icon" href="/icon.svg" />
        <meta name="apple-mobile-web-app-capable" content="yes" />
        <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent" />
        <meta name="apple-mobile-web-app-title" content="Трекер" />
      </head>
      <body>{children}</body>
    </html>
  )
}
