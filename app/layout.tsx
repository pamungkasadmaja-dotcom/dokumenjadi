import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: {
    default: "Dokumen Jadi - Buat Surat, CV, Invoice, dan Notulen Online",
    template: "%s | Dokumen Jadi",
  },
  description:
    "Buat dokumen formal online dengan mudah. Tersedia template surat izin, surat pengunduran diri, surat lamaran kerja, CV, invoice, notulen rapat, berita acara, surat kuasa, dan surat pernyataan.",
  verification: {
  google: "rXPW-YCmzFOCu1XfFSpOYqxBBEzrAdeSOpJYfg0Yd5g",
},
    keywords: [
    "buat surat online",
    "template surat",
    "surat izin",
    "surat pengunduran diri",
    "surat lamaran kerja",
    "buat CV online",
    "invoice online",
    "notulen rapat",
    "berita acara",
    "surat kuasa",
    "surat pernyataan",
  ],
  authors: [{ name: "Dokumen Jadi" }],
  creator: "Dokumen Jadi",
  publisher: "Dokumen Jadi",
  robots: {
    index: true,
    follow: true,
  },
  openGraph: {
    title: "Dokumen Jadi - Buat Dokumen Formal Online",
    description:
      "Buat surat, CV, invoice, notulen rapat, berita acara, dan dokumen formal lain secara online. Bisa download Word dan PDF.",
    url: "https://dokumenjadi.vercel.app",
    siteName: "Dokumen Jadi",
    locale: "id_ID",
    type: "website",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html
      lang="id"
      className={`${geistSans.variable} ${geistMono.variable} h-full antialiased`}
    >
      <body className="min-h-full flex flex-col">{children}</body>
    </html>
  );
}
