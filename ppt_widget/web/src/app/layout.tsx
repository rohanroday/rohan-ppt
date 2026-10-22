import type { Metadata } from "next";
import Script from "next/script";
import "./globals.css";

export const metadata: Metadata = {
  title: "Presentation Builder",
  description: "Generate PowerPoint presentations from Zoho CRM records",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="en" className="h-full antialiased">
      {/* Browser extensions (e.g. ColorZilla) add attributes to <body> before React loads. */}
      <body className="min-h-full flex flex-col" suppressHydrationWarning>
        {/* Vendored Zoho widget SDK, served from this package (no CDN at runtime). */}
        <Script src="./vendor/ZohoEmbededAppSDK.min.js" strategy="beforeInteractive" />
        {children}
      </body>
    </html>
  );
}
