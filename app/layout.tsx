import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'Translation and Evaluation Form Extraction',
  description: 'Parse Outcomes translation table pastes and Evaluation Form captures into editable exports.',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  )
}


