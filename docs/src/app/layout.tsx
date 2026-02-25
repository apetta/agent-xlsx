import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";
import { ThemeProvider } from "@/components/theme-provider";
import { MobileNavProvider } from "@/components/mobile-nav-context";
import { Header } from "@/components/header";
import { Sidebar } from "@/components/sidebar";

const geist = Geist({
  variable: "--font-geist",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "agent-xlsx",
  description: "Excel file CLI built with Agent Experience (AX) in mind.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" suppressHydrationWarning>
      <head>
        <meta name="theme-color" content="#0f0f0f" media="(prefers-color-scheme: dark)" />
        <meta name="theme-color" content="#ffffff" media="(prefers-color-scheme: light)" />
      </head>
      <body
        className={`${geist.variable} ${geistMono.variable} antialiased bg-background text-foreground`}
        suppressHydrationWarning
      >
        <ThemeProvider>
          <MobileNavProvider>
            <a
              href="#main-content"
              className="sr-only focus:not-sr-only focus:fixed focus:top-2 focus:left-2 focus:z-[100] focus:px-4 focus:py-2 focus:bg-background focus:text-foreground focus:border focus:border-border focus:rounded-md focus:text-sm"
            >
              Skip to content
            </a>
            <Header />
            <div className="flex min-h-[calc(100vh-3.5rem)]">
              <Sidebar />
              <main id="main-content" className="flex-1 overflow-auto">
                <div className="max-w-2xl mx-auto px-4 sm:px-6 py-8 sm:py-12">
                  <div className="prose">
                    {children}
                  </div>
                </div>
              </main>
            </div>
          </MobileNavProvider>
        </ThemeProvider>
      </body>
    </html>
  );
}
