import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Template Surat Izin Gratis - Buat Surat Izin Online | DokumenJadi",
  description:
    "Buat template surat izin online gratis untuk izin tidak masuk kerja, izin sekolah, izin sakit, dan keperluan lainnya. Praktis, rapi, dan bisa diunduh.",
  alternates: {
    canonical: "https://dokumenjadi.vercel.app/template-surat-izin",
  },
  openGraph: {
    title: "Template Surat Izin Gratis - DokumenJadi",
    description:
      "Buat surat izin online secara cepat dan rapi. Cocok untuk izin kerja, sekolah, sakit, dan keperluan lainnya.",
    url: "https://dokumenjadi.vercel.app/template-surat-izin",
    siteName: "DokumenJadi",
    type: "website",
  },
};

export default function TemplateSuratIzinPage() {
  const jsonLd = {
    "@context": "https://schema.org",
    "@type": "WebPage",
    name: "Template Surat Izin Gratis",
    description:
      "Halaman untuk membuat template surat izin online gratis di DokumenJadi.",
    url: "https://dokumenjadi.vercel.app/template-surat-izin",
    publisher: {
      "@type": "Organization",
      name: "DokumenJadi",
    },
  };

  return (
    <main className="mx-auto max-w-3xl px-6 py-12">
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLd) }}
      />

      <h1 className="text-3xl font-bold mb-4">
        Template Surat Izin Gratis
      </h1>

      <p className="mb-4 text-gray-700">
        Gunakan DokumenJadi untuk membuat surat izin secara online dengan cepat,
        rapi, dan mudah diunduh. Template ini dapat digunakan untuk surat izin
        tidak masuk kerja, surat izin sekolah, surat izin sakit, dan kebutuhan
        administrasi lainnya.
      </p>

      <h2 className="text-2xl font-semibold mt-8 mb-3">
        Jenis Surat Izin yang Bisa Dibuat
      </h2>

      <ul className="list-disc pl-6 space-y-2 text-gray-700">
        <li>Surat izin tidak masuk kerja</li>
        <li>Surat izin tidak masuk sekolah</li>
        <li>Surat izin sakit</li>
        <li>Surat izin keperluan keluarga</li>
        <li>Surat izin acara atau kegiatan tertentu</li>
      </ul>

      <h2 className="text-2xl font-semibold mt-8 mb-3">
        Mengapa Menggunakan DokumenJadi?
      </h2>

      <p className="mb-4 text-gray-700">
        DokumenJadi membantu pengguna membuat dokumen formal tanpa harus menyusun
        format dari awal. Pengguna cukup mengisi data yang diperlukan, kemudian
        dokumen dapat dibuat dalam format yang rapi dan siap digunakan.
      </p>

      <h2 className="text-2xl font-semibold mt-8 mb-3">
        Contoh Format Surat Izin
      </h2>

      <div className="rounded-xl border p-5 bg-gray-50 text-gray-700">
        <p>Yth. Bapak/Ibu [Nama Penerima]</p>
        <p className="mt-4">
          Dengan hormat, saya yang bertanda tangan di bawah ini menyampaikan
          permohonan izin untuk tidak dapat hadir pada [tanggal] karena
          [alasan izin].
        </p>
        <p className="mt-4">
          Demikian surat izin ini dibuat. Atas perhatian dan pengertiannya,
          saya ucapkan terima kasih.
        </p>
      </div>

      <h2 className="text-2xl font-semibold mt-8 mb-3">
        Cara Membuat Surat Izin
      </h2>

      <ol className="list-decimal pl-6 space-y-2 text-gray-700">
        <li>Pilih jenis dokumen surat izin.</li>
        <li>Isi data pemohon, tujuan surat, tanggal, dan alasan izin.</li>
        <li>Periksa kembali isi surat.</li>
        <li>Unduh dokumen dalam format yang tersedia.</li>
      </ol>

      <a
        href="/"
        className="inline-block mt-8 rounded-xl bg-blue-600 px-5 py-3 text-white font-semibold"
      >
        Buat Surat Izin Sekarang
      </a>

      <h2 className="text-2xl font-semibold mt-10 mb-3">
        Pertanyaan Umum
      </h2>

      <h3 className="font-semibold mt-4">
        Apakah template surat izin ini gratis?
      </h3>
      <p className="text-gray-700">
        Ya, pengguna dapat membuat surat izin secara online melalui DokumenJadi.
      </p>

      <h3 className="font-semibold mt-4">
        Apakah bisa digunakan untuk surat izin kerja?
      </h3>
      <p className="text-gray-700">
        Bisa. Template surat izin dapat disesuaikan untuk kebutuhan izin kerja,
        sekolah, sakit, atau keperluan lainnya.
      </p>
    </main>
  );
}