"use client";

import { useRef, useState } from "react";
import type { FormEvent, PointerEvent } from "react";
import {
  AlignmentType,
  BorderStyle,
  Document,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

import { jsPDF } from "jspdf";

type Field = {
  name: string;
  label: string;
  type: "text" | "textarea" | "date" | "email" | "number";
  required: boolean;
};

type Template = {
  id: string;
  name: string;
  description: string;
  fields: Field[];
};

type NotulenEntry = {
  pembahasan: string;
  keputusan: string;
  tindakLanjut: string;
};

type InvoiceItem = {
  deskripsi: string;
  jumlah: string;
  hargaSatuan: string;
};

type UiLanguage = "id" | "en";

const templates: Template[] = [
  {
    id: "surat-izin",
    name: "Surat Izin",
    description: "Untuk izin tidak masuk kerja, sekolah, kuliah, atau kegiatan.",
    fields: [
      { name: "nama", label: "Nama Lengkap", type: "text", required: true },
      { name: "alamat", label: "Alamat", type: "textarea", required: true },
      { name: "tujuanSurat", label: "Tujuan Surat", type: "text", required: true },
      { name: "tanggalMulaiIzin", label: "Tanggal Mulai Izin", type: "date", required: true },
{ name: "tanggalSelesaiIzin", label: "Tanggal Selesai Izin", type: "date", required: true },
      { name: "alasan", label: "Alasan Izin", type: "textarea", required: true },
      { name: "kota", label: "Kota Pembuatan Surat", type: "text", required: true },
      { name: "tanggalSurat", label: "Tanggal Surat", type: "date", required: true },
    ],
  },
  {
    id: "surat-pengunduran-diri",
    name: "Surat Pengunduran Diri",
    description: "Untuk membuat surat resign yang sopan dan profesional.",
    fields: [
      { name: "nama", label: "Nama Lengkap", type: "text", required: true },
      { name: "jabatan", label: "Jabatan", type: "text", required: true },
      { name: "perusahaan", label: "Nama Perusahaan", type: "text", required: true },
      { name: "tanggalTerakhir", label: "Tanggal Terakhir Bekerja", type: "date", required: true },
      { name: "alasan", label: "Alasan Pengunduran Diri", type: "textarea", required: false },
      { name: "kota", label: "Kota Pembuatan Surat", type: "text", required: true },
      { name: "tanggalSurat", label: "Tanggal Surat", type: "date", required: true },
    ],
  },
  {
    id: "surat-lamaran-kerja",
    name: "Surat Lamaran Kerja",
    description: "Untuk membuat surat lamaran kerja yang rapi dan siap digunakan.",
    fields: [
      { name: "nama", label: "Nama Lengkap", type: "text", required: true },
      { name: "alamat", label: "Alamat", type: "textarea", required: true },
      { name: "email", label: "Email", type: "email", required: true },
      { name: "telepon", label: "Nomor HP", type: "text", required: true },
      { name: "perusahaan", label: "Nama Perusahaan Tujuan", type: "text", required: true },
      { name: "posisi", label: "Posisi yang Dilamar", type: "text", required: true },
      { name: "pendidikan", label: "Pendidikan Terakhir", type: "text", required: true },
      { name: "pengalaman", label: "Pengalaman Singkat", type: "textarea", required: false },
      { name: "kota", label: "Kota Pembuatan Surat", type: "text", required: true },
      { name: "tanggalSurat", label: "Tanggal Surat", type: "date", required: true },
    ],
  },
  {
    id: "surat-kuasa",
    name: "Surat Kuasa",
    description: "Untuk memberikan kuasa kepada pihak lain dalam suatu keperluan.",
    fields: [
      { name: "namaPemberi", label: "Nama Pemberi Kuasa", type: "text", required: true },
      { name: "nikPemberi", label: "NIK Pemberi Kuasa", type: "text", required: true },
      { name: "alamatPemberi", label: "Alamat Pemberi Kuasa", type: "textarea", required: true },
      { name: "namaPenerima", label: "Nama Penerima Kuasa", type: "text", required: true },
      { name: "nikPenerima", label: "NIK Penerima Kuasa", type: "text", required: true },
      { name: "alamatPenerima", label: "Alamat Penerima Kuasa", type: "textarea", required: true },
      { name: "keperluan", label: "Keperluan Kuasa", type: "textarea", required: true },
      { name: "kota", label: "Kota Pembuatan Surat", type: "text", required: true },
      { name: "tanggalSurat", label: "Tanggal Surat", type: "date", required: true },
    ],
  },
  {
    id: "surat-pernyataan",
    name: "Surat Pernyataan",
    description: "Untuk membuat surat pernyataan pribadi atau formal.",
    fields: [
      { name: "nama", label: "Nama Lengkap", type: "text", required: true },
      { name: "nik", label: "NIK", type: "text", required: false },
      { name: "alamat", label: "Alamat", type: "textarea", required: true },
      { name: "isiPernyataan", label: "Isi Pernyataan", type: "textarea", required: true },
      { name: "kota", label: "Kota Pembuatan Surat", type: "text", required: true },
      { name: "tanggalSurat", label: "Tanggal Surat", type: "date", required: true },
    ],
  },
  {
    id: "berita-acara",
    name: "Berita Acara",
    description: "Untuk membuat berita acara kegiatan, serah terima, rapat, atau kejadian.",
    fields: [
  { name: "judul", label: "Judul / Perihal Berita Acara", type: "text", required: true },
  { name: "nomorBeritaAcara", label: "Nomor Berita Acara (Opsional)", type: "text", required: false },
  { name: "hariTanggal", label: "Hari / Tanggal", type: "text", required: true },
  { name: "waktu", label: "Waktu", type: "text", required: true },
  { name: "tempat", label: "Tempat", type: "text", required: true },
  { name: "peserta", label: "Pihak / Peserta yang Terlibat (Opsional)", type: "textarea", required: false },
  { name: "uraian", label: "Uraian Berita Acara", type: "textarea", required: true },
  { name: "namaPenandatangan", label: "Nama Penandatangan", type: "text", required: true },
  { name: "kota", label: "Kota Pembuatan Dokumen", type: "text", required: true },
  { name: "tanggalSurat", label: "Tanggal Dokumen", type: "date", required: true },
],
  },
  {
    id: "notulen-rapat",
    name: "Notulen Rapat",
    description: "Untuk membuat notulen rapat yang rapi dan mudah dibagikan.",
   fields: [
  { name: "judulRapat", label: "Judul / Nama Rapat", type: "text", required: true },
  { name: "tanggalRapat", label: "Tanggal Rapat", type: "date", required: true },
  { name: "waktuRapat", label: "Waktu Rapat", type: "text", required: true },
  { name: "tempatRapat", label: "Tempat / Media Rapat", type: "text", required: true },
  { name: "pimpinanRapat", label: "Pimpinan Rapat", type: "text", required: true },
  { name: "notulis", label: "Notulis", type: "text", required: true },
  { name: "pesertaRapat", label: "Peserta Rapat", type: "textarea", required: true },
  { name: "agenda", label: "Agenda Rapat", type: "textarea", required: true },
  { name: "kota", label: "Kota Pembuatan Notulen", type: "text", required: true },
  { name: "tanggalDokumen", label: "Tanggal Dokumen", type: "date", required: true },
],
  },
  {
    id: "proposal-kegiatan",
    name: "Proposal Kegiatan",
    description: "Untuk membuat proposal kegiatan sederhana dan terstruktur.",
    fields: [
      { name: "namaKegiatan", label: "Nama Kegiatan", type: "text", required: true },
      { name: "latarBelakang", label: "Latar Belakang", type: "textarea", required: true },
      { name: "tujuan", label: "Tujuan Kegiatan", type: "textarea", required: true },
      { name: "waktuKegiatan", label: "Waktu Kegiatan", type: "text", required: true },
      { name: "tempatKegiatan", label: "Tempat Kegiatan", type: "text", required: true },
      { name: "sasaranPeserta", label: "Sasaran Peserta", type: "textarea", required: true },
      { name: "anggaran", label: "Rencana Anggaran", type: "textarea", required: false },
      { name: "namaPenanggungJawab", label: "Nama Penanggung Jawab", type: "text", required: true },
      { name: "kota", label: "Kota Pembuatan Proposal", type: "text", required: true },
      { name: "tanggalDokumen", label: "Tanggal Proposal", type: "date", required: true },
    ],
  },
  {
    id: "invoice",
    name: "Invoice",
    description: "Untuk membuat invoice sederhana bagi UMKM, freelancer, atau jasa.",
    fields: [
      { name: "namaBisnis", label: "Nama Bisnis / Penjual", type: "text", required: true },
      { name: "namaPelanggan", label: "Nama Pelanggan", type: "text", required: true },
      { name: "nomorInvoice", label: "Nomor Invoice", type: "text", required: true },
      { name: "tanggalInvoice", label: "Tanggal Invoice", type: "date", required: true },
      { name: "catatan", label: "Catatan Pembayaran", type: "textarea", required: false },
    ],
  },
  {
    id: "cv-sederhana",
    name: "CV Sederhana",
    description: "Untuk membuat CV sederhana yang rapi dan mudah dibaca.",
    fields: [
      { name: "nama", label: "Nama Lengkap", type: "text", required: true },
      { name: "profesi", label: "Profesi / Posisi yang Diinginkan", type: "text", required: true },
      { name: "email", label: "Email", type: "email", required: true },
      { name: "telepon", label: "Nomor HP", type: "text", required: true },
      { name: "ringkasan", label: "Ringkasan Profil", type: "textarea", required: true },
      { name: "pendidikan", label: "Riwayat Pendidikan", type: "textarea", required: true },
      { name: "pengalaman", label: "Pengalaman Kerja", type: "textarea", required: false },
      { name: "keahlian", label: "Keahlian", type: "textarea", required: true },
    ],
  },
];

function formatTanggalIndonesia(tanggal: string) {
  if (!tanggal) return "";

  return new Date(tanggal).toLocaleDateString("id-ID", {
    day: "numeric",
    month: "long",
    year: "numeric",
  });
}

function formatTanggalInggris(tanggal: string) {
  if (!tanggal) return "";

  return new Date(tanggal).toLocaleDateString("en-US", {
    day: "numeric",
    month: "long",
    year: "numeric",
  });
}

const currencyOptions = [
  { value: "IDR", label: "Rupiah Indonesia (IDR)" },
  { value: "USD", label: "US Dollar (USD)" },
  { value: "EUR", label: "Euro (EUR)" },
  { value: "SGD", label: "Singapore Dollar (SGD)" },
  { value: "MYR", label: "Malaysian Ringgit (MYR)" },
  { value: "JPY", label: "Japanese Yen (JPY)" },
  { value: "SAR", label: "Saudi Riyal (SAR)" },
  { value: "AUD", label: "Australian Dollar (AUD)" },
];

const uiText: Record<UiLanguage, Record<string, string>> = {
  id: {
    heroTitle: "Buat dokumen otomatis dalam hitungan detik.",
    heroDescription:
      "Pilih template, isi form sederhana, lalu download dokumen yang rapi dan siap digunakan.",
    chooseTemplate: "Pilih Template",
    formHelper: "Isi data berikut untuk membuat dokumen.",

    discussionListTitle: "Daftar Pembahasan Rapat",
    discussionListDescription:
      "Tambahkan pembahasan, keputusan, dan tindak lanjut sesuai kebutuhan.",
    addEntry: "+ Tambah Entry",
    entry: "Entry",
    remove: "Hapus",
    discussion: "Pembahasan",
    decision: "Keputusan",
    followUp: "Tindak Lanjut",

    currency: "Mata Uang",
    invoiceItemListTitle: "Daftar Item Tagihan",
    invoiceItemListDescription:
      "Tambahkan item, jumlah, dan harga satuan sesuai kebutuhan invoice.",
    addItem: "+ Tambah Item",
    item: "Item",
    productDescription: "Deskripsi Produk/Jasa",
    quantity: "Jumlah",
    unitPrice: "Harga Satuan",

    documentLanguage: "Bahasa Dokumen",
    indonesian: "Indonesia",
    english: "Inggris",

    signature: "Tanda Tangan",
    signatureInstruction:
      "Buat tanda tangan langsung pada area di atas, lalu pastikan posisinya berada di tengah area.",
    clearSignature: "Hapus Tanda Tangan",

    clearInputs: "Kosongkan Isian",
    downloadWord: "Download Word",
    downloadPdf: "Download PDF",

    signatureRequired: "Tanda tangan wajib diisi.",
    wordUnavailable: "Download Word belum tersedia untuk template ini.",
    pdfUnavailable: "Download PDF belum tersedia untuk template ini.",
    endDateBeforeStart:
      "Tanggal selesai izin tidak boleh lebih awal dari tanggal mulai izin.",
    discussionEntryRequired: "Minimal isi 1 entry pembahasan rapat.",
    discussionColumnRequired:
      "Kolom Pembahasan wajib diisi pada setiap entry yang digunakan.",
    invoiceItemRequired: "Minimal isi 1 item tagihan.",
    invoiceDescriptionRequired:
      "Deskripsi wajib diisi pada setiap item tagihan.",
    invoiceAmountRequired:
      "Jumlah dan harga satuan pada setiap item harus lebih dari 0.",

    clearModalTitle: "Kosongkan Semua Isian?",
    clearModalBody:
      "Semua data yang sudah diisi, termasuk tanda tangan, akan dikosongkan. Apakah Anda yakin ingin melanjutkan?",
    no: "Tidak",
    yesClear: "Ya, Kosongkan",
  },

  en: {
    heroTitle: "Create documents automatically in seconds.",
    heroDescription:
      "Choose a template, fill in a simple form, then download a neat and ready-to-use document.",
    chooseTemplate: "Choose Template",
    formHelper: "Fill in the information below to create your document.",

    discussionListTitle: "Meeting Discussion List",
    discussionListDescription:
      "Add discussions, decisions, and follow-up actions as needed.",
    addEntry: "+ Add Entry",
    entry: "Entry",
    remove: "Remove",
    discussion: "Discussion",
    decision: "Decision",
    followUp: "Follow-up",

    currency: "Currency",
    invoiceItemListTitle: "Billing Item List",
    invoiceItemListDescription:
      "Add items, quantities, and unit prices as needed for the invoice.",
    addItem: "+ Add Item",
    item: "Item",
    productDescription: "Product/Service Description",
    quantity: "Quantity",
    unitPrice: "Unit Price",

    documentLanguage: "Document Language",
    indonesian: "Indonesian",
    english: "English",

    signature: "Signature",
    signatureInstruction:
      "Write your signature directly in the area above and make sure it is positioned in the center.",
    clearSignature: "Clear Signature",

    clearInputs: "Clear Inputs",
    downloadWord: "Download Word",
    downloadPdf: "Download PDF",

    signatureRequired: "Signature is required.",
    wordUnavailable: "Word download is not available for this template yet.",
    pdfUnavailable: "PDF download is not available for this template yet.",
    endDateBeforeStart: "The end date cannot be earlier than the start date.",
    discussionEntryRequired: "Please fill in at least 1 meeting discussion entry.",
    discussionColumnRequired:
      "The Discussion field is required for every used entry.",
    invoiceItemRequired: "Please fill in at least 1 billing item.",
    invoiceDescriptionRequired:
      "Description is required for every billing item.",
    invoiceAmountRequired:
      "Quantity and unit price for every item must be greater than 0.",

    clearModalTitle: "Clear All Inputs?",
    clearModalBody:
      "All entered data, including the signature, will be cleared. Are you sure you want to continue?",
    no: "No",
    yesClear: "Yes, Clear",
  },
};

const templateText: Record<
  string,
  Record<UiLanguage, { name: string; description: string }>
> = {
  "surat-izin": {
    id: {
      name: "Surat Izin",
      description: "Untuk izin tidak masuk kerja, sekolah, kuliah, atau kegiatan.",
    },
    en: {
      name: "Permission Letter",
      description: "For requesting leave from work, school, college, or activities.",
    },
  },
  "surat-pengunduran-diri": {
    id: {
      name: "Surat Pengunduran Diri",
      description: "Untuk membuat surat resign yang sopan dan profesional.",
    },
    en: {
      name: "Resignation Letter",
      description: "For creating a polite and professional resignation letter.",
    },
  },
  "surat-lamaran-kerja": {
    id: {
      name: "Surat Lamaran Kerja",
      description: "Untuk membuat surat lamaran kerja yang rapi dan siap digunakan.",
    },
    en: {
      name: "Job Application Letter",
      description: "For creating a neat and ready-to-use job application letter.",
    },
  },
  "surat-kuasa": {
    id: {
      name: "Surat Kuasa",
      description: "Untuk memberikan kuasa kepada pihak lain dalam suatu keperluan.",
    },
    en: {
      name: "Power of Attorney",
      description: "For authorizing another person to act on your behalf.",
    },
  },
  "surat-pernyataan": {
    id: {
      name: "Surat Pernyataan",
      description: "Untuk membuat surat pernyataan pribadi atau formal.",
    },
    en: {
      name: "Statement Letter",
      description: "For creating a personal or formal statement letter.",
    },
  },
  "berita-acara": {
    id: {
      name: "Berita Acara",
      description: "Untuk membuat berita acara kegiatan, serah terima, rapat, atau kejadian.",
    },
    en: {
      name: "Official Report",
      description: "For creating reports of activities, handovers, meetings, or events.",
    },
  },
  "notulen-rapat": {
    id: {
      name: "Notulen Rapat",
      description: "Untuk membuat notulen rapat yang rapi dan mudah dibagikan.",
    },
    en: {
      name: "Meeting Minutes",
      description: "For creating neat and shareable meeting minutes.",
    },
  },
  "proposal-kegiatan": {
    id: {
      name: "Proposal Kegiatan",
      description: "Untuk membuat proposal kegiatan sederhana dan terstruktur.",
    },
    en: {
      name: "Activity Proposal",
      description: "For creating a simple and structured activity proposal.",
    },
  },
  invoice: {
    id: {
      name: "Invoice",
      description: "Untuk membuat invoice sederhana bagi UMKM, freelancer, atau jasa.",
    },
    en: {
      name: "Invoice",
      description: "For creating a simple invoice for small businesses, freelancers, or services.",
    },
  },
  "cv-sederhana": {
    id: {
      name: "Curriculum Vitae",
      description: "Untuk membuat daftar riwayat hidup yang rapi dan mudah dibaca.",
    },
    en: {
      name: "Curriculum Vitae",
      description: "For creating a neat and easy-to-read curriculum vitae.",
    },
  },
};

const fieldLabelText: Record<string, Record<UiLanguage, string>> = {
  nama: { id: "Nama Lengkap", en: "Full Name" },
  alamat: { id: "Alamat", en: "Address" },
  tujuanSurat: { id: "Tujuan Surat", en: "Letter Recipient / Purpose" },
  tanggalMulaiIzin: { id: "Tanggal Mulai Izin", en: "Leave Start Date" },
  tanggalSelesaiIzin: { id: "Tanggal Selesai Izin", en: "Leave End Date" },
  alasan: { id: "Alasan", en: "Reason" },
  kota: { id: "Kota Pembuatan Dokumen", en: "Document City" },
  tanggalSurat: { id: "Tanggal Surat", en: "Letter Date" },

  jabatan: { id: "Jabatan", en: "Position" },
  perusahaan: { id: "Nama Perusahaan", en: "Company Name" },
  tanggalTerakhir: { id: "Tanggal Terakhir Bekerja", en: "Last Working Date" },

  email: { id: "Email", en: "Email" },
  telepon: { id: "Nomor HP", en: "Phone Number" },
  posisi: { id: "Posisi yang Dilamar", en: "Position Applied For" },
  pendidikan: { id: "Pendidikan Terakhir", en: "Latest Education" },
  pengalaman: { id: "Pengalaman Singkat", en: "Brief Experience" },

  namaPemberi: { id: "Nama Pemberi Kuasa", en: "Grantor Name" },
  nikPemberi: { id: "NIK Pemberi Kuasa", en: "Grantor ID Number" },
  alamatPemberi: { id: "Alamat Pemberi Kuasa", en: "Grantor Address" },
  namaPenerima: { id: "Nama Penerima Kuasa", en: "Attorney Name" },
  nikPenerima: { id: "NIK Penerima Kuasa", en: "Attorney ID Number" },
  alamatPenerima: { id: "Alamat Penerima Kuasa", en: "Attorney Address" },
  keperluan: { id: "Keperluan Kuasa", en: "Authorization Purpose" },

  nik: { id: "NIK", en: "ID Number" },
  isiPernyataan: { id: "Isi Pernyataan", en: "Statement Content" },

  judul: { id: "Judul / Perihal Berita Acara", en: "Official Report Title / Subject" },
  nomorBeritaAcara: { id: "Nomor Berita Acara (Opsional)", en: "Official Report Number (Optional)" },
  hariTanggal: { id: "Hari / Tanggal", en: "Day / Date" },
  waktu: { id: "Waktu", en: "Time" },
  tempat: { id: "Tempat", en: "Place" },
  peserta: { id: "Pihak / Peserta yang Terlibat (Opsional)", en: "Related Parties / Participants (Optional)" },
  uraian: { id: "Uraian Berita Acara", en: "Report Description" },
  namaPenandatangan: { id: "Nama Penandatangan", en: "Signer Name" },

  judulRapat: { id: "Judul / Nama Rapat", en: "Meeting Title / Name" },
  tanggalRapat: { id: "Tanggal Rapat", en: "Meeting Date" },
  waktuRapat: { id: "Waktu Rapat", en: "Meeting Time" },
  tempatRapat: { id: "Tempat / Media Rapat", en: "Meeting Place / Media" },
  pimpinanRapat: { id: "Pimpinan Rapat", en: "Meeting Chair" },
  notulis: { id: "Notulis", en: "Minutes Prepared By" },
  pesertaRapat: { id: "Peserta Rapat", en: "Participants" },
  agenda: { id: "Agenda Rapat", en: "Meeting Agenda" },
  tanggalDokumen: { id: "Tanggal Dokumen", en: "Document Date" },

  namaKegiatan: { id: "Nama Kegiatan", en: "Activity Name" },
  latarBelakang: { id: "Latar Belakang", en: "Background" },
  tujuan: { id: "Tujuan Kegiatan", en: "Activity Objectives" },
  waktuKegiatan: { id: "Waktu Kegiatan", en: "Activity Time" },
  tempatKegiatan: { id: "Tempat Kegiatan", en: "Activity Location" },
  sasaranPeserta: { id: "Sasaran Peserta", en: "Target Participants" },
  anggaran: { id: "Rencana Anggaran", en: "Budget Plan" },
  namaPenanggungJawab: { id: "Nama Penanggung Jawab", en: "Person in Charge" },

  namaBisnis: { id: "Nama Bisnis / Penjual", en: "Business / Seller Name" },
  namaPelanggan: { id: "Nama Pelanggan", en: "Customer Name" },
  nomorInvoice: { id: "Nomor Invoice", en: "Invoice Number" },
  tanggalInvoice: { id: "Tanggal Invoice", en: "Invoice Date" },
  catatan: { id: "Catatan Pembayaran", en: "Payment Notes" },

  profesi: { id: "Profesi / Posisi yang Diinginkan", en: "Profession / Target Position" },
  ringkasan: { id: "Ringkasan Profil", en: "Profile Summary" },
  keahlian: { id: "Keahlian", en: "Skills" },
};

function getCurrencyLabel(currency: string) {
  return (
    currencyOptions.find((item) => item.value === currency)?.label || currency
  );
}

function getCurrencyLocale(currency: string, isEnglish: boolean) {
  if (currency === "IDR") return "id-ID";
  if (currency === "JPY") return "ja-JP";
  if (currency === "SAR") return "ar-SA";
  if (currency === "MYR") return "ms-MY";
  if (currency === "SGD") return "en-SG";
  if (currency === "AUD") return "en-AU";
  if (currency === "EUR") return "de-DE";
  if (currency === "USD") return "en-US";

  return isEnglish ? "en-US" : "id-ID";
}

function formatCurrencyAmount(
  angka: number,
  currency: string,
  isEnglish: boolean
) {
  const noDecimalCurrencies = ["IDR", "JPY", "KRW", "VND"];

  const fractionDigits = noDecimalCurrencies.includes(currency) ? 0 : 2;

  return new Intl.NumberFormat(getCurrencyLocale(currency, isEnglish), {
    style: "currency",
    currency,
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits,
  }).format(angka);
}

function dataUrlToUint8Array(dataUrl: string) {
  const base64 = dataUrl.split(",")[1];
  const binaryString = window.atob(base64);
  const bytes = new Uint8Array(binaryString.length);

  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }

  return bytes;
}

function getCenteredSignatureDataUrl(canvas: HTMLCanvasElement) {
  const context = canvas.getContext("2d");
  if (!context) return canvas.toDataURL("image/png");

  const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
  const data = imageData.data;

  let minX = canvas.width;
  let minY = canvas.height;
  let maxX = 0;
  let maxY = 0;
  let hasDrawing = false;

  for (let y = 0; y < canvas.height; y++) {
    for (let x = 0; x < canvas.width; x++) {
      const index = (y * canvas.width + x) * 4;
      const alpha = data[index + 3];

      if (alpha > 0) {
        hasDrawing = true;
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x);
        maxY = Math.max(maxY, y);
      }
    }
  }

  if (!hasDrawing) return "";

  const signatureWidth = maxX - minX;
  const signatureHeight = maxY - minY;

  const outputCanvas = document.createElement("canvas");
  outputCanvas.width = 500;
  outputCanvas.height = 220;

  const outputContext = outputCanvas.getContext("2d");
  if (!outputContext) return canvas.toDataURL("image/png");

  const padding = 40;
  const maxOutputWidth = outputCanvas.width - padding * 2;
  const maxOutputHeight = outputCanvas.height - padding * 2;

  const scale = Math.min(
    maxOutputWidth / signatureWidth,
    maxOutputHeight / signatureHeight
  );

  const drawWidth = signatureWidth * scale;
  const drawHeight = signatureHeight * scale;

  const drawX = (outputCanvas.width - drawWidth) / 2;
  const drawY = (outputCanvas.height - drawHeight) / 2;

  outputContext.drawImage(
    canvas,
    minX,
    minY,
    signatureWidth,
    signatureHeight,
    drawX,
    drawY,
    drawWidth,
    drawHeight
  );

  return outputCanvas.toDataURL("image/png");
}

export default function Home() {
  const [selectedTemplateId, setSelectedTemplateId] = useState("surat-izin");
  const [uiLanguage, setUiLanguage] = useState<UiLanguage>("id");

const signatureCanvasRef = useRef<HTMLCanvasElement | null>(null);
const formRef = useRef<HTMLFormElement | null>(null);
const isDrawingSignature = useRef(false);
const [signatureData, setSignatureData] = useState("");
const [showClearConfirm, setShowClearConfirm] = useState(false);

const [notulenEntries, setNotulenEntries] = useState<NotulenEntry[]>([
  {
    pembahasan: "",
    keputusan: "",
    tindakLanjut: "",
  },
]);

function addNotulenEntry() {
  setNotulenEntries((previousEntries) => [
    ...previousEntries,
    {
      pembahasan: "",
      keputusan: "",
      tindakLanjut: "",
    },
  ]);
}

function updateNotulenEntry(
  index: number,
  field: keyof NotulenEntry,
  value: string
) {
  setNotulenEntries((previousEntries) =>
    previousEntries.map((entry, entryIndex) =>
      entryIndex === index
        ? {
            ...entry,
            [field]: value,
          }
        : entry
    )
  );
}

function removeNotulenEntry(index: number) {
  setNotulenEntries((previousEntries) => {
    if (previousEntries.length === 1) {
      return previousEntries;
    }

    return previousEntries.filter((_, entryIndex) => entryIndex !== index);
  });
}

const [invoiceItems, setInvoiceItems] = useState<InvoiceItem[]>([
  {
    deskripsi: "",
    jumlah: "",
    hargaSatuan: "",
  },
]);

function addInvoiceItem() {
  setInvoiceItems((previousItems) => [
    ...previousItems,
    {
      deskripsi: "",
      jumlah: "",
      hargaSatuan: "",
    },
  ]);
}

function updateInvoiceItem(
  index: number,
  field: keyof InvoiceItem,
  value: string
) {
  setInvoiceItems((previousItems) =>
    previousItems.map((item, itemIndex) =>
      itemIndex === index
        ? {
            ...item,
            [field]: value,
          }
        : item
    )
  );
}

function removeInvoiceItem(index: number) {
  setInvoiceItems((previousItems) => {
    if (previousItems.length === 1) {
      return previousItems;
    }

    return previousItems.filter((_, itemIndex) => itemIndex !== index);
  });
}

  const selectedTemplate = templates.find(
    (template) => template.id === selectedTemplateId
  );

const t = uiText[uiLanguage];

function getTemplateName(template: Template) {
  return templateText[template.id]?.[uiLanguage]?.name || template.name;
}

function getTemplateDescription(template: Template) {
  return (
    templateText[template.id]?.[uiLanguage]?.description ||
    template.description
  );
}

function getFieldLabel(field: Field) {
  return fieldLabelText[field.name]?.[uiLanguage] || field.label;
}

function handleDownloadPdf(
  values: Record<string, FormDataEntryValue>,
  isEnglish: boolean
) {
  const pdf = new jsPDF("p", "mm", "a4");

  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();

  const marginLeft = 25;
  const marginRight = 25;
  const marginTop = 25;
  const marginBottom = 22;
  const contentWidth = pageWidth - marginLeft - marginRight;

  let y = marginTop;

  type PdfAlign = "left" | "center" | "right";

  const getValue = (key: string) => String(values[key] || "").trim();

  const safeFileName = (text: string) =>
    text.replace(/[\\/:*?"<>|]/g, "-").trim() || "DokumenJadi.pdf";

  const formatDate = (tanggal: string) =>
    isEnglish ? formatTanggalInggris(tanggal) : formatTanggalIndonesia(tanggal);

  const toPdfFileName = (fileName: string) =>
    safeFileName(fileName.replace(/\.docx$/i, ".pdf"));

  function checkPage(requiredSpace = 12) {
    if (y + requiredSpace > pageHeight - marginBottom) {
      pdf.addPage();
      y = marginTop;
    }
  }

  function setFont(size = 12, bold = false) {
    pdf.setFont("times", bold ? "bold" : "normal");
    pdf.setFontSize(size);
  }

  function splitLines(text: string, width: number) {
    const paragraphs = (text || "-").split("\n");
    const result: string[] = [];

    paragraphs.forEach((paragraph) => {
      const wrapped = pdf.splitTextToSize(
        paragraph || " ",
        width
      ) as string[];

      result.push(...wrapped);
    });

    return result.length > 0 ? result : [" "];
  }

  function addText(
    text: string,
    x: number,
    width: number,
    options: {
      align?: PdfAlign;
      size?: number;
      bold?: boolean;
      after?: number;
      lineHeight?: number;
    } = {}
  ) {
    const align = options.align || "left";
    const size = options.size || 12;
    const bold = options.bold || false;
    const after = options.after ?? 4;
    const lineHeight = options.lineHeight || 6.5;

    setFont(size, bold);

    const lines = splitLines(text || "-", width);

    lines.forEach((line) => {
      checkPage(lineHeight + 2);

      let textX = x;

      if (align === "center") {
        textX = x + width / 2;
      }

      if (align === "right") {
        textX = x + width;
      }

      pdf.text(line, textX, y, { align });
      y += lineHeight;
    });

    y += after;
  }

  function addCenteredText(text: string, size = 14, bold = false, after = 8) {
    setFont(size, bold);
    checkPage(10);
    pdf.text(text, pageWidth / 2, y, { align: "center" });
    y += after;
  }

  function addRightText(text: string, size = 12, after = 8) {
    setFont(size, false);
    checkPage(10);
    pdf.text(text, pageWidth - marginRight, y, { align: "right" });
    y += after;
  }

  function addParagraph(text: string, after = 7) {
    addText(text, marginLeft, contentWidth, {
      size: 12,
      after,
      lineHeight: 6.5,
    });
  }

  function addSectionTitle(text: string) {
    y += 4;
    checkPage(12);
    setFont(13, true);
    pdf.text(text, marginLeft, y);
    y += 8;
  }

  type PdfTableCell = {
    text: string;
    width: number;
    bold?: boolean;
    align?: PdfAlign;
    size?: number;
  };

  function addTable(
    rows: PdfTableCell[][],
    options: {
      x?: number;
      after?: number;
      lineHeight?: number;
      minRowHeight?: number;
    } = {}
  ) {
    const tableX = options.x ?? marginLeft;
    const after = options.after ?? 6;
    const lineHeight = options.lineHeight ?? 5.5;
    const minRowHeight = options.minRowHeight ?? 11;
    const paddingX = 3;
    const paddingTop = 4;

    rows.forEach((row) => {
      const cellLines = row.map((cell) => {
        setFont(cell.size || 11, cell.bold || false);
        return splitLines(cell.text || " ", cell.width - paddingX * 2);
      });

      const maxLines = Math.max(...cellLines.map((lines) => lines.length));
      const rowHeight = Math.max(
        minRowHeight,
        maxLines * lineHeight + paddingTop + 4
      );

      checkPage(rowHeight + 3);

      let x = tableX;

      row.forEach((cell, cellIndex) => {
        pdf.setDrawColor(0);
        pdf.setLineWidth(0.2);
        pdf.rect(x, y, cell.width, rowHeight);

        setFont(cell.size || 11, cell.bold || false);

        const align = cell.align || "left";
        const lines = cellLines[cellIndex];

        lines.forEach((line, lineIndex) => {
          const textY = y + paddingTop + 4 + lineIndex * lineHeight;

          if (align === "right") {
            pdf.text(line, x + cell.width - paddingX, textY, {
              align: "right",
            });
          } else if (align === "center") {
            pdf.text(line, x + cell.width / 2, textY, {
              align: "center",
            });
          } else {
            pdf.text(line, x + paddingX, textY);
          }
        });

        x += cell.width;
      });

      y += rowHeight;
    });

    y += after;
  }

  function addSignatureBlock({
    city,
    dateText,
    roleText,
    signerName,
    minY = 0,
    compact = false,
  }: {
    city?: string;
    dateText?: string;
    roleText: string;
    signerName: string;
    minY?: number;
    compact?: boolean;
  }) {
    if (minY > 0 && y < minY) {
      y = minY;
    }

    const hasCityOrDate = Boolean(city || dateText);
    const neededSpace = compact
      ? hasCityOrDate
        ? 40
        : 34
      : hasCityOrDate
        ? 58
        : 46;

    checkPage(neededSpace);

    const blockWidth = 65;
    const blockX = pageWidth - marginRight - blockWidth;
    const centerX = blockX + blockWidth / 2;

    const dateGap = compact ? 8 : 12;
    const roleGap = compact ? 5 : 6;
    const signatureWidth = compact ? 30 : 34;
    const signatureHeight = compact ? 16 : 20;
    const afterSignature = compact ? 21 : 28;
    const afterName = compact ? 6 : 8;

    setFont(12, false);

    if (hasCityOrDate) {
      const cityDateText =
        city && dateText ? `${city}, ${dateText}` : city || dateText || "";

      pdf.text(cityDateText, centerX, y, { align: "center" });
      y += dateGap;
    }

    pdf.text(roleText, centerX, y, { align: "center" });
    y += roleGap;

    if (signatureData) {
      pdf.addImage(
        signatureData,
        "PNG",
        centerX - signatureWidth / 2,
        y,
        signatureWidth,
        signatureHeight
      );
    }

    y += afterSignature;

    pdf.text(signerName || "-", centerX, y, { align: "center" });
    y += afterName;
  }

  function renderSimpleLetter({
    title,
    city,
    dateText,
    destination,
    identity1,
    identity2,
    body,
    closing,
    signerName,
    fileName,
  }: {
    title: string;
    city: string;
    dateText: string;
    destination: string;
    identity1: string;
    identity2: string;
    body: string;
    closing: string;
    signerName: string;
    fileName: string;
  }) {
    y = marginTop;

    addCenteredText(title, 16, true, 14);
    addRightText(`${city}, ${dateText}`, 12, 14);

    addParagraph(isEnglish ? "To:" : "Kepada Yth.", 1);
    addParagraph(destination, 10);

    addParagraph(isEnglish ? "Respectfully," : "Dengan hormat,", 12);

    addParagraph(
      isEnglish
        ? "I, the undersigned below:"
        : "Saya yang bertanda tangan di bawah ini:",
      10
    );

    addParagraph(identity1, 1);
    addParagraph(identity2, 10);

    addParagraph(body, 10);
    addParagraph(closing, 8);

    addSignatureBlock({
      roleText: isEnglish ? "Sincerely," : "Hormat saya,",
      signerName,
      minY: 205,
    });

    pdf.save(toPdfFileName(fileName));
  }

  if (selectedTemplateId === "surat-izin") {
    const nama = getValue("nama");
    const alamat = getValue("alamat");
    const tujuanSurat = getValue("tujuanSurat");
    const tanggalMulai = getValue("tanggalMulaiIzin");
    const tanggalSelesai = getValue("tanggalSelesaiIzin");
    const alasan = getValue("alasan");
    const kota = getValue("kota");
    const tanggalSurat = getValue("tanggalSurat");

    if (tanggalSelesai < tanggalMulai) {
      alert(t.endDateBeforeStart);
      return;
    }

    const tanggalMulaiFormatted = formatDate(tanggalMulai);
    const tanggalSelesaiFormatted = formatDate(tanggalSelesai);
    const tanggalSuratFormatted = formatDate(tanggalSurat);

    renderSimpleLetter({
      title: isEnglish ? "PERMISSION LETTER" : "SURAT IZIN",
      city: kota,
      dateText: tanggalSuratFormatted,
      destination: tujuanSurat,
      identity1: isEnglish ? `Name    : ${nama}` : `Nama    : ${nama}`,
      identity2: isEnglish ? `Address : ${alamat}` : `Alamat  : ${alamat}`,
      body: isEnglish
        ? `I hereby request permission to be absent from ${tanggalMulaiFormatted} to ${tanggalSelesaiFormatted} due to ${alasan}.`
        : `Dengan ini saya mengajukan izin tidak masuk pada tanggal ${tanggalMulaiFormatted} sampai dengan ${tanggalSelesaiFormatted} karena ${alasan}.`,
      closing: isEnglish
        ? "Thus, I submit this permission letter truthfully. Thank you for your attention and consideration."
        : "Demikian surat izin ini saya buat dengan sebenar-benarnya. Atas perhatian dan kebijaksanaannya, saya ucapkan terima kasih.",
      signerName: nama,
      fileName: isEnglish
        ? `Permission Letter - ${nama}.docx`
        : `Surat Izin - ${nama}.docx`,
    });

    return;
  }

  if (selectedTemplateId === "surat-pengunduran-diri") {
    const nama = getValue("nama");
    const jabatan = getValue("jabatan");
    const perusahaan = getValue("perusahaan");
    const tanggalTerakhir = getValue("tanggalTerakhir");
    const alasan = getValue("alasan");
    const kota = getValue("kota");
    const tanggalSurat = getValue("tanggalSurat");

    const tanggalSuratFormatted = formatDate(tanggalSurat);
    const tanggalTerakhirFormatted = formatDate(tanggalTerakhir);

    renderSimpleLetter({
      title: isEnglish ? "RESIGNATION LETTER" : "SURAT PENGUNDURAN DIRI",
      city: kota,
      dateText: tanggalSuratFormatted,
      destination: perusahaan,
      identity1: isEnglish ? `Name     : ${nama}` : `Nama     : ${nama}`,
      identity2: isEnglish ? `Position : ${jabatan}` : `Jabatan  : ${jabatan}`,
      body: isEnglish
        ? `I hereby submit my resignation from my position as ${jabatan} at ${perusahaan}, effective ${tanggalTerakhirFormatted}.${alasan ? ` The reason for my resignation is ${alasan}.` : ""}`
        : `Dengan ini saya mengajukan pengunduran diri dari jabatan saya sebagai ${jabatan} di ${perusahaan}, terhitung efektif pada tanggal ${tanggalTerakhirFormatted}.${alasan ? ` Adapun alasan pengunduran diri saya adalah ${alasan}.` : ""}`,
      closing: isEnglish
        ? "I would like to express my gratitude for the opportunity, trust, and experience given during my employment. I hope the company continues to grow and succeed."
        : "Saya mengucapkan terima kasih atas kesempatan, kepercayaan, dan pengalaman yang telah diberikan selama saya bekerja. Semoga perusahaan terus berkembang dan semakin sukses.",
      signerName: nama,
      fileName: isEnglish
        ? `Resignation Letter - ${nama}.docx`
        : `Surat Pengunduran Diri - ${nama}.docx`,
    });

    return;
  }

  if (selectedTemplateId === "surat-lamaran-kerja") {
    const nama = getValue("nama");
    const alamat = getValue("alamat");
    const email = getValue("email");
    const telepon = getValue("telepon");
    const perusahaan = getValue("perusahaan");
    const posisi = getValue("posisi");
    const pendidikan = getValue("pendidikan");
    const pengalaman = getValue("pengalaman");
    const kota = getValue("kota");
    const tanggalSurat = getValue("tanggalSurat");

    const tanggalSuratFormatted = formatDate(tanggalSurat);

    renderSimpleLetter({
      title: isEnglish ? "JOB APPLICATION LETTER" : "SURAT LAMARAN KERJA",
      city: kota,
      dateText: tanggalSuratFormatted,
      destination: isEnglish
        ? `HR Department of ${perusahaan}`
        : `Yth. HRD ${perusahaan}`,
      identity1: isEnglish ? `Name    : ${nama}` : `Nama    : ${nama}`,
      identity2: isEnglish ? `Address : ${alamat}` : `Alamat  : ${alamat}`,
      body: isEnglish
        ? `I hereby submit my application for the position of ${posisi} at ${perusahaan}. I have an educational background in ${pendidikan} and relevant experience in ${pengalaman}. I can be contacted by email at ${email} or by phone at ${telepon}.`
        : `Dengan ini saya mengajukan lamaran kerja untuk posisi ${posisi} di ${perusahaan}. Saya memiliki latar belakang pendidikan ${pendidikan} serta pengalaman singkat di bidang ${pengalaman}. Saya dapat dihubungi melalui email ${email} atau nomor telepon ${telepon}.`,
      closing: isEnglish
        ? "I hope to be given the opportunity to proceed to the next stage of the recruitment process. Thank you for your attention and consideration."
        : "Besar harapan saya untuk diberikan kesempatan mengikuti tahapan seleksi berikutnya. Atas perhatian dan kesempatan yang diberikan, saya ucapkan terima kasih.",
      signerName: nama,
      fileName: isEnglish
        ? `Job Application Letter - ${nama}.docx`
        : `Surat Lamaran Kerja - ${nama}.docx`,
    });

    return;
  }

  if (selectedTemplateId === "surat-kuasa") {
    const namaPemberi = getValue("namaPemberi");
    const nikPemberi = getValue("nikPemberi");
    const alamatPemberi = getValue("alamatPemberi");
    const namaPenerima = getValue("namaPenerima");
    const nikPenerima = getValue("nikPenerima");
    const alamatPenerima = getValue("alamatPenerima");
    const keperluan = getValue("keperluan");
    const kota = getValue("kota");
    const tanggalSurat = getValue("tanggalSurat");

    const tanggalSuratFormatted = formatDate(tanggalSurat);

    renderSimpleLetter({
      title: isEnglish ? "POWER OF ATTORNEY" : "SURAT KUASA",
      city: kota,
      dateText: tanggalSuratFormatted,
      destination: isEnglish
        ? "To whom it may concern"
        : "Pihak yang berkepentingan",
      identity1: isEnglish
        ? `Grantor : ${namaPemberi}, ID Number ${nikPemberi}, Address ${alamatPemberi}`
        : `Pemberi Kuasa : ${namaPemberi}, NIK ${nikPemberi}, Alamat ${alamatPemberi}`,
      identity2: isEnglish
        ? `Attorney : ${namaPenerima}, ID Number ${nikPenerima}, Address ${alamatPenerima}`
        : `Penerima Kuasa : ${namaPenerima}, NIK ${nikPenerima}, Alamat ${alamatPenerima}`,
      body: isEnglish
        ? `I, as the grantor, hereby authorize the attorney named above to act on my behalf for the following purpose: ${keperluan}. This power of attorney is made to be used properly and responsibly.`
        : `Saya selaku pemberi kuasa dengan ini memberikan kuasa kepada penerima kuasa tersebut di atas untuk bertindak atas nama saya dalam keperluan: ${keperluan}. Surat kuasa ini dibuat untuk dipergunakan sebagaimana mestinya dan dengan penuh tanggung jawab.`,
      closing: isEnglish
        ? "Thus, this power of attorney is made truthfully and may be used as necessary."
        : "Demikian surat kuasa ini dibuat dengan sebenar-benarnya untuk dipergunakan sebagaimana mestinya.",
      signerName: namaPemberi,
      fileName: isEnglish
        ? `Power of Attorney - ${namaPemberi}.docx`
        : `Surat Kuasa - ${namaPemberi}.docx`,
    });

    return;
  }

  if (selectedTemplateId === "surat-pernyataan") {
    const nama = getValue("nama");
    const nik = getValue("nik");
    const alamat = getValue("alamat");
    const isiPernyataan = getValue("isiPernyataan");
    const kota = getValue("kota");
    const tanggalSurat = getValue("tanggalSurat");

    const tanggalSuratFormatted = formatDate(tanggalSurat);

    renderSimpleLetter({
      title: isEnglish ? "STATEMENT LETTER" : "SURAT PERNYATAAN",
      city: kota,
      dateText: tanggalSuratFormatted,
      destination: isEnglish
        ? "To whom it may concern"
        : "Pihak yang berkepentingan",
      identity1: isEnglish
        ? `Name : ${nama}, ID Number ${nik}`
        : `Nama : ${nama}, NIK ${nik}`,
      identity2: isEnglish ? `Address : ${alamat}` : `Alamat : ${alamat}`,
      body: isEnglish
        ? `I hereby declare that ${isiPernyataan}. This statement is made truthfully and can be used as necessary.`
        : `Dengan ini saya menyatakan bahwa ${isiPernyataan}. Pernyataan ini saya buat dengan sebenar-benarnya untuk dipergunakan sebagaimana mestinya.`,
      closing: isEnglish
        ? "Thus, this statement letter is made truthfully and without coercion from any party."
        : "Demikian surat pernyataan ini saya buat dengan sebenar-benarnya dan tanpa paksaan dari pihak mana pun.",
      signerName: nama,
      fileName: isEnglish
        ? `Statement Letter - ${nama}.docx`
        : `Surat Pernyataan - ${nama}.docx`,
    });

    return;
  }

  if (selectedTemplateId === "berita-acara") {
    const judulBeritaAcara = getValue("judul");
    const nomorBeritaAcara = getValue("nomorBeritaAcara");
    const hariTanggal = getValue("hariTanggal");
    const waktu = getValue("waktu");
    const tempat = getValue("tempat");
    const peserta = getValue("peserta");
    const uraian = getValue("uraian");
    const namaPenandatangan = getValue("namaPenandatangan");
    const kota = getValue("kota");
    const tanggalSurat = getValue("tanggalSurat");

    const tanggalDokumenFormatted = formatDate(tanggalSurat);

    y = marginTop;

    addCenteredText(isEnglish ? "OFFICIAL REPORT" : "BERITA ACARA", 16, true, 8);
    addCenteredText(judulBeritaAcara.toUpperCase(), 13, true, nomorBeritaAcara ? 8 : 14);

    if (nomorBeritaAcara) {
      addCenteredText(
        isEnglish ? `Number: ${nomorBeritaAcara}` : `Nomor: ${nomorBeritaAcara}`,
        12,
        false,
        14
      );
    }

    addParagraph(
      isEnglish
        ? `On ${hariTanggal}, at ${waktu}, located at ${tempat}, this official report was prepared regarding ${judulBeritaAcara}.`
        : `Pada hari ini, ${hariTanggal}, pukul ${waktu}, bertempat di ${tempat}, telah dibuat Berita Acara mengenai ${judulBeritaAcara}.`,
      8
    );

    addParagraph(
      isEnglish
        ? "The details are as follows:"
        : "Adapun rincian berita acara adalah sebagai berikut:",
      5
    );

    const labelWidth = contentWidth * 0.3;
    const valueWidth = contentWidth * 0.7;

    addTable([
      [
        {
          text: isEnglish ? "Subject" : "Judul / Perihal",
          width: labelWidth,
          bold: true,
        },
        { text: judulBeritaAcara, width: valueWidth },
      ],
      [
        {
          text: isEnglish ? "Day / Date" : "Hari / Tanggal",
          width: labelWidth,
          bold: true,
        },
        { text: hariTanggal, width: valueWidth },
      ],
      [
        { text: isEnglish ? "Time" : "Waktu", width: labelWidth, bold: true },
        { text: waktu, width: valueWidth },
      ],
      [
        { text: isEnglish ? "Place" : "Tempat", width: labelWidth, bold: true },
        { text: tempat, width: valueWidth },
      ],
      ...(peserta
        ? [
            [
              {
                text: isEnglish
                  ? "Related Parties / Participants"
                  : "Pihak / Peserta yang Terlibat",
                width: labelWidth,
                bold: true,
              },
              { text: peserta, width: valueWidth },
            ],
          ]
        : []),
      [
        {
          text: isEnglish ? "Description" : "Uraian",
          width: labelWidth,
          bold: true,
        },
        { text: uraian, width: valueWidth },
      ],
    ]);

    addParagraph(
      isEnglish
        ? "Thus, this official report is made truthfully as documentation and may be used as necessary."
        : "Demikian Berita Acara ini dibuat dengan sebenar-benarnya sebagai bukti/dokumentasi dan untuk dipergunakan sebagaimana mestinya.",
      8
    );

    addSignatureBlock({
      city: kota,
      dateText: tanggalDokumenFormatted,
      roleText: isEnglish
        ? "Prepared and signed by,"
        : "Pihak yang Menyusun Berita Acara,",
      signerName: namaPenandatangan,
    });

    pdf.save(
      safeFileName(
        isEnglish
          ? `Official Report - ${judulBeritaAcara}.pdf`
          : `Berita Acara - ${judulBeritaAcara}.pdf`
      )
    );

    return;
  }

  if (selectedTemplateId === "notulen-rapat") {
    const judulRapat = getValue("judulRapat");
    const tanggalRapat = getValue("tanggalRapat");
    const waktuRapat = getValue("waktuRapat");
    const tempatRapat = getValue("tempatRapat");
    const pimpinanRapat = getValue("pimpinanRapat");
    const notulis = getValue("notulis");
    const pesertaRapat = getValue("pesertaRapat");
    const agenda = getValue("agenda");
    const kota = getValue("kota");
    const tanggalDokumen = getValue("tanggalDokumen");

    const pembahasanRows = notulenEntries
      .map((entry) => ({
        pembahasan: entry.pembahasan.trim(),
        keputusan: entry.keputusan.trim(),
        tindakLanjut: entry.tindakLanjut.trim(),
      }))
      .filter(
        (entry) =>
          entry.pembahasan !== "" ||
          entry.keputusan !== "" ||
          entry.tindakLanjut !== ""
      );

    if (pembahasanRows.length === 0) {
      alert(t.discussionEntryRequired);
      return;
    }

    if (pembahasanRows.some((entry) => entry.pembahasan === "")) {
      alert(t.discussionColumnRequired);
      return;
    }

    const tanggalRapatFormatted = formatDate(tanggalRapat);
    const tanggalDokumenFormatted = formatDate(tanggalDokumen);

    y = marginTop;

    addCenteredText(isEnglish ? "MEETING MINUTES" : "NOTULEN RAPAT", 16, true, 8);
    addCenteredText(judulRapat.toUpperCase(), 13, true, 14);

    const labelWidth = contentWidth * 0.3;
    const valueWidth = contentWidth * 0.7;

    addTable([
      [
        {
          text: isEnglish ? "Meeting Title" : "Judul / Nama Rapat",
          width: labelWidth,
          bold: true,
        },
        { text: judulRapat, width: valueWidth },
      ],
      [
        {
          text: isEnglish ? "Date" : "Tanggal Rapat",
          width: labelWidth,
          bold: true,
        },
        { text: tanggalRapatFormatted, width: valueWidth },
      ],
      [
        {
          text: isEnglish ? "Time" : "Waktu Rapat",
          width: labelWidth,
          bold: true,
        },
        { text: waktuRapat, width: valueWidth },
      ],
      [
        {
          text: isEnglish ? "Place / Media" : "Tempat / Media Rapat",
          width: labelWidth,
          bold: true,
        },
        { text: tempatRapat, width: valueWidth },
      ],
      [
        {
          text: isEnglish ? "Meeting Chair" : "Pimpinan Rapat",
          width: labelWidth,
          bold: true,
        },
        { text: pimpinanRapat, width: valueWidth },
      ],
      [
        {
          text: isEnglish ? "Minutes Prepared By" : "Notulis",
          width: labelWidth,
          bold: true,
        },
        { text: notulis, width: valueWidth },
      ],
      [
        {
          text: isEnglish ? "Participants" : "Peserta Rapat",
          width: labelWidth,
          bold: true,
        },
        { text: pesertaRapat, width: valueWidth },
      ],
      [
        { text: isEnglish ? "Agenda" : "Agenda Rapat", width: labelWidth, bold: true },
        { text: agenda, width: valueWidth },
      ],
    ]);

    addSectionTitle(
      isEnglish
        ? "Discussion, Decisions, and Follow-up Actions"
        : "Pembahasan, Keputusan, dan Tindak Lanjut"
    );

    addTable(
      [
        [
          { text: "No.", width: contentWidth * 0.08, bold: true, align: "center" },
          {
            text: isEnglish ? "Discussion" : "Pembahasan",
            width: contentWidth * 0.37,
            bold: true,
            align: "center",
          },
          {
            text: isEnglish ? "Decision" : "Keputusan",
            width: contentWidth * 0.27,
            bold: true,
            align: "center",
          },
          {
            text: isEnglish ? "Follow-up" : "Tindak Lanjut",
            width: contentWidth * 0.28,
            bold: true,
            align: "center",
          },
        ],
        ...pembahasanRows.map((entry, index) => [
          {
            text: String(index + 1),
            width: contentWidth * 0.08,
            align: "center" as PdfAlign,
          },
          { text: entry.pembahasan, width: contentWidth * 0.37 },
          { text: entry.keputusan || "-", width: contentWidth * 0.27 },
          { text: entry.tindakLanjut || "-", width: contentWidth * 0.28 },
        ]),
      ],
      { minRowHeight: 12 }
    );

    addParagraph(
      isEnglish
        ? "These meeting minutes are prepared as official documentation of the discussion, decisions, and follow-up actions agreed upon in the meeting."
        : "Demikian notulen rapat ini disusun sebagai dokumentasi resmi atas pembahasan, keputusan, dan tindak lanjut yang disepakati dalam rapat.",
      8
    );

    addSignatureBlock({
      city: kota,
      dateText: tanggalDokumenFormatted,
      roleText: isEnglish ? "Minutes prepared by," : "Notulis,",
      signerName: notulis,
    });

    pdf.save(
      safeFileName(
        isEnglish
          ? `Meeting Minutes - ${judulRapat}.pdf`
          : `Notulen Rapat - ${judulRapat}.pdf`
      )
    );

    return;
  }

  if (selectedTemplateId === "proposal-kegiatan") {
    const namaKegiatan = getValue("namaKegiatan");
    const latarBelakang = getValue("latarBelakang");
    const tujuan = getValue("tujuan");
    const waktuKegiatan = getValue("waktuKegiatan");
    const tempatKegiatan = getValue("tempatKegiatan");
    const sasaranPeserta = getValue("sasaranPeserta");
    const anggaran = getValue("anggaran");
    const namaPenanggungJawab = getValue("namaPenanggungJawab");
    const kota = getValue("kota");
    const tanggalDokumen = getValue("tanggalDokumen");

    const tanggalDokumenFormatted = formatDate(tanggalDokumen);

    y = marginTop;

    addCenteredText(
      isEnglish ? "ACTIVITY PROPOSAL" : "PROPOSAL KEGIATAN",
      16,
      true,
      8
    );

    addCenteredText(namaKegiatan.toUpperCase(), 14, true, 12);

    addSectionTitle(isEnglish ? "A. Background" : "A. Latar Belakang");
    addParagraph(latarBelakang);

    addSectionTitle(
      isEnglish ? "B. Purpose and Objectives" : "B. Maksud dan Tujuan"
    );
    addParagraph(tujuan);

    addSectionTitle(
      isEnglish
        ? "C. Time and Place of Activity"
        : "C. Waktu dan Tempat Kegiatan"
    );

    addTable([
      [
        { text: isEnglish ? "Time" : "Waktu", width: contentWidth * 0.3, bold: true },
        { text: waktuKegiatan, width: contentWidth * 0.7 },
      ],
      [
        { text: isEnglish ? "Place" : "Tempat", width: contentWidth * 0.3, bold: true },
        { text: tempatKegiatan, width: contentWidth * 0.7 },
      ],
    ]);

    addSectionTitle(isEnglish ? "D. Target Participants" : "D. Sasaran Peserta");
    addParagraph(sasaranPeserta);

    addSectionTitle(isEnglish ? "E. Budget Plan" : "E. Rencana Anggaran");
    addParagraph(anggaran || "-");

    addSectionTitle(isEnglish ? "F. Closing" : "F. Penutup");
    addParagraph(
      isEnglish
        ? "Thus, this activity proposal is prepared to serve as a reference for the implementation of the activity. Thank you for your attention and support."
        : "Demikian proposal kegiatan ini disusun sebagai acuan pelaksanaan kegiatan. Atas perhatian dan dukungan yang diberikan, kami ucapkan terima kasih.",
      4
    );

    addSignatureBlock({
      city: kota,
      dateText: tanggalDokumenFormatted,
      roleText: isEnglish ? "Person in charge," : "Penanggung Jawab,",
      signerName: namaPenanggungJawab,
      compact: true,
    });

    pdf.save(
      safeFileName(
        isEnglish
          ? `Activity Proposal - ${namaKegiatan}.pdf`
          : `Proposal Kegiatan - ${namaKegiatan}.pdf`
      )
    );

    return;
  }

  if (selectedTemplateId === "invoice") {
    const namaBisnis = getValue("namaBisnis");
const namaPelanggan = getValue("namaPelanggan");
const nomorInvoice = getValue("nomorInvoice");
const tanggalInvoice = getValue("tanggalInvoice");
const mataUang = getValue("mataUang") || "IDR";
const catatan = getValue("catatan");

    const invoiceRows = invoiceItems
      .map((item) => {
        const jumlah = Number(item.jumlah || 0);
        const hargaSatuan = Number(item.hargaSatuan || 0);

        return {
          deskripsi: item.deskripsi.trim(),
          jumlah,
          hargaSatuan,
          total: jumlah * hargaSatuan,
        };
      })
      .filter(
        (item) =>
          item.deskripsi !== "" || item.jumlah > 0 || item.hargaSatuan > 0
      );

    if (invoiceRows.length === 0) {
      alert(t.invoiceItemRequired);
      return;
    }

    if (invoiceRows.some((item) => item.deskripsi === "")) {
      alert(t.invoiceDescriptionRequired);
      return;
    }

    if (
      invoiceRows.some(
        (item) =>
          !Number.isFinite(item.jumlah) ||
          !Number.isFinite(item.hargaSatuan) ||
          item.jumlah <= 0 ||
          item.hargaSatuan <= 0
      )
    ) {
      alert(t.invoiceAmountRequired);
      return;
    }

    const tanggalInvoiceFormatted = formatDate(tanggalInvoice);

    const formatUang = (angka: number) =>
  formatCurrencyAmount(angka, mataUang, isEnglish);

    const formatAngka = (angka: number) =>
      new Intl.NumberFormat(isEnglish ? "en-US" : "id-ID", {
        maximumFractionDigits: 2,
      }).format(angka);

    const totalTagihan = invoiceRows.reduce(
      (sum, item) => sum + item.total,
      0
    );

    y = 35;

    addCenteredText("INVOICE", 18, true, 8);
    addCenteredText(nomorInvoice, 12, false, 14);

    addTable([
      [
        {
          text: isEnglish ? "Seller / Business" : "Penjual / Bisnis",
          width: contentWidth * 0.3,
          bold: true,
        },
        { text: namaBisnis, width: contentWidth * 0.7 },
      ],
      [
        {
          text: isEnglish ? "Customer" : "Pelanggan",
          width: contentWidth * 0.3,
          bold: true,
        },
        { text: namaPelanggan, width: contentWidth * 0.7 },
      ],
      [
  {
    text: isEnglish ? "Currency" : "Mata Uang",
    width: contentWidth * 0.3,
    bold: true,
  },
  { text: getCurrencyLabel(mataUang), width: contentWidth * 0.7 },
],
    ]);

    addSectionTitle(isEnglish ? "Billing Details" : "Rincian Tagihan");

    addTable([
      [
        {
          text: isEnglish ? "Description" : "Deskripsi",
          width: contentWidth * 0.4,
          bold: true,
        },
        {
          text: isEnglish ? "Qty" : "Jumlah",
          width: contentWidth * 0.15,
          bold: true,
          align: "center",
        },
        {
          text: isEnglish ? "Unit Price" : "Harga Satuan",
          width: contentWidth * 0.22,
          bold: true,
        },
        {
          text: "Total",
          width: contentWidth * 0.23,
          bold: true,
        },
      ],
      ...invoiceRows.map((item) => [
        {
          text: item.deskripsi,
          width: contentWidth * 0.4,
        },
        {
          text: formatAngka(item.jumlah),
          width: contentWidth * 0.15,
          align: "center" as PdfAlign,
        },
        {
          text: formatUang(item.hargaSatuan),
          width: contentWidth * 0.22,
        },
        {
          text: formatUang(item.total),
          width: contentWidth * 0.23,
        },
      ]),
      [
        { text: " ", width: contentWidth * 0.4 },
        { text: " ", width: contentWidth * 0.15 },
        {
          text: isEnglish ? "Grand Total" : "Total Tagihan",
          width: contentWidth * 0.22,
          bold: true,
        },
        {
          text: formatUang(totalTagihan),
          width: contentWidth * 0.23,
          bold: true,
        },
      ],
    ]);

    if (catatan) {
      addSectionTitle(isEnglish ? "Payment Notes" : "Catatan Pembayaran");
      addParagraph(catatan);
    }

    addParagraph(
      isEnglish
        ? "Thank you for your trust and cooperation."
        : "Terima kasih atas kepercayaan dan kerja samanya.",
      8
    );

    addSignatureBlock({
      roleText: isEnglish ? "Seller," : "Penjual,",
      signerName: namaBisnis,
      minY: 205,
    });

    pdf.save(safeFileName(`Invoice - ${nomorInvoice}.pdf`));
    return;
  }

  if (selectedTemplateId === "cv-sederhana") {
    const nama = getValue("nama");
    const profesi = getValue("profesi");
    const email = getValue("email");
    const telepon = getValue("telepon");
    const ringkasan = getValue("ringkasan");
    const pendidikan = getValue("pendidikan");
    const pengalaman = getValue("pengalaman");
    const keahlian = getValue("keahlian");

    y = marginTop;

    addCenteredText(nama.toUpperCase(), 18, true, 8);
    addCenteredText(profesi, 13, false, 8);
    addCenteredText("CURRICULUM VITAE", 14, true, 14);

    addTable([
      [
        { text: `Email: ${email}`, width: contentWidth * 0.5, bold: false },
        {
          text: `${isEnglish ? "Phone" : "Nomor HP"}: ${telepon}`,
          width: contentWidth * 0.5,
          bold: false,
        },
      ],
    ]);

    addSectionTitle(isEnglish ? "Profile Summary" : "Ringkasan Profil");
    addParagraph(ringkasan);

    addSectionTitle(isEnglish ? "Education" : "Riwayat Pendidikan");
    addParagraph(pendidikan);

    addSectionTitle(isEnglish ? "Work Experience" : "Pengalaman Kerja");
    addParagraph(pengalaman || "-");

    addSectionTitle(isEnglish ? "Skills" : "Keahlian");
    addParagraph(keahlian);

    addSignatureBlock({
      roleText: isEnglish ? "Prepared by," : "Dibuat oleh,",
      signerName: nama,
    });

    pdf.save(safeFileName(`Curriculum Vitae - ${nama}.pdf`));
    return;
  }

  alert(t.pdfUnavailable);
}

function getCanvasPoint(event: PointerEvent<HTMLCanvasElement>) {
  const canvas = signatureCanvasRef.current;

  if (!canvas) {
    return { x: 0, y: 0 };
  }

  const rect = canvas.getBoundingClientRect();

  return {
    x: ((event.clientX - rect.left) * canvas.width) / rect.width,
    y: ((event.clientY - rect.top) * canvas.height) / rect.height,
  };
}

function startSignature(event: PointerEvent<HTMLCanvasElement>) {
  const canvas = signatureCanvasRef.current;
  if (!canvas) return;

  const context = canvas.getContext("2d");
  if (!context) return;

  event.currentTarget.setPointerCapture(event.pointerId);

  isDrawingSignature.current = true;

  const point = getCanvasPoint(event);

  context.beginPath();
  context.moveTo(point.x, point.y);
}

function drawSignature(event: PointerEvent<HTMLCanvasElement>) {
  if (!isDrawingSignature.current) return;

  const canvas = signatureCanvasRef.current;
  if (!canvas) return;

  const context = canvas.getContext("2d");
  if (!context) return;

  event.preventDefault();

  const point = getCanvasPoint(event);

  context.lineWidth = 3;
  context.lineCap = "round";
  context.lineJoin = "round";
  context.strokeStyle = "#000000";

  context.lineTo(point.x, point.y);
  context.stroke();
}

function endSignature() {
  const canvas = signatureCanvasRef.current;
  if (!canvas) return;

  isDrawingSignature.current = false;

  const centeredSignature = getCenteredSignatureDataUrl(canvas);
  setSignatureData(centeredSignature);
}

function clearSignature() {
  const canvas = signatureCanvasRef.current;
  if (!canvas) return;

  const context = canvas.getContext("2d");
  if (!context) return;

  context.clearRect(0, 0, canvas.width, canvas.height);
  setSignatureData("");
}

function clearAllInputs() {
  const form = formRef.current;

  if (form) {
    form
      .querySelectorAll<HTMLInputElement | HTMLTextAreaElement>(
        "input, textarea"
      )
      .forEach((element) => {
        element.value = "";
      });

    const languageSelect = form.querySelector<HTMLSelectElement>(
      'select[name="bahasaDokumen"]'
    );

    if (languageSelect) {
      languageSelect.value = "id";
    }

    const currencySelect = form.querySelector<HTMLSelectElement>(
  'select[name="mataUang"]'
);

if (currencySelect) {
  currencySelect.value = "IDR";
}
  }


  
  clearSignature();

  setNotulenEntries([
  {
    pembahasan: "",
    keputusan: "",
    tindakLanjut: "",
  },
]);

  setInvoiceItems([
    {
      deskripsi: "",
      jumlah: "",
      hargaSatuan: "",
    },
  ]);

  setShowClearConfirm(false);
}

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
  event.preventDefault();

  const submitter = (event.nativeEvent as SubmitEvent)
    .submitter as HTMLButtonElement | null;

  const downloadType = submitter?.value || "word";

  const formData = new FormData(event.currentTarget);
  const values = Object.fromEntries(formData.entries()) as Record<
    string,
    FormDataEntryValue
  >;

if (!signatureData) {
  alert(t.signatureRequired);
  return;
}

if (
  selectedTemplateId !== "surat-izin" &&
  selectedTemplateId !== "surat-pengunduran-diri" &&
  selectedTemplateId !== "surat-lamaran-kerja" &&
  selectedTemplateId !== "surat-kuasa" &&
  selectedTemplateId !== "surat-pernyataan" &&
  selectedTemplateId !== "berita-acara" &&
  selectedTemplateId !== "notulen-rapat" &&
  selectedTemplateId !== "proposal-kegiatan" &&
  selectedTemplateId !== "invoice" &&
  selectedTemplateId !== "cv-sederhana"
) {
  alert(t.wordUnavailable);
  return;
}



  const bahasaDokumen = values.bahasaDokumen as string;
const isEnglish = bahasaDokumen === "en";

if (downloadType === "pdf") {
  handleDownloadPdf(values, isEnglish);
  return;
}

let nama = "";
let kota = "";
let tanggalSuratFormatted = "";
let judul = "";
let tujuanSurat = "";
let identitasBaris1 = "";
let identitasBaris2 = "";
let isiSurat = "";
let penutup = "";
let namaFile = "";

if (selectedTemplateId === "surat-izin") {
  nama = values.nama as string;
  const alamat = values.alamat as string;
  tujuanSurat = values.tujuanSurat as string;
  const tanggalMulai = values.tanggalMulaiIzin as string;
  const tanggalSelesai = values.tanggalSelesaiIzin as string;
  const alasan = values.alasan as string;
  kota = values.kota as string;
  const tanggalSurat = values.tanggalSurat as string;

  if (tanggalSelesai < tanggalMulai) {
    alert(t.endDateBeforeStart);
    return;
  }

  tanggalSuratFormatted = isEnglish
    ? formatTanggalInggris(tanggalSurat)
    : formatTanggalIndonesia(tanggalSurat);

  const tanggalMulaiFormatted = isEnglish
    ? formatTanggalInggris(tanggalMulai)
    : formatTanggalIndonesia(tanggalMulai);

  const tanggalSelesaiFormatted = isEnglish
    ? formatTanggalInggris(tanggalSelesai)
    : formatTanggalIndonesia(tanggalSelesai);

  judul = isEnglish ? "PERMISSION LETTER" : "SURAT IZIN";

  identitasBaris1 = isEnglish
    ? `Name    : ${nama}`
    : `Nama    : ${nama}`;

  identitasBaris2 = isEnglish
    ? `Address : ${alamat}`
    : `Alamat  : ${alamat}`;

  isiSurat = isEnglish
    ? `I hereby request permission to be absent from ${tanggalMulaiFormatted} to ${tanggalSelesaiFormatted} due to ${alasan}.`
    : `Dengan ini saya mengajukan izin tidak masuk pada tanggal ${tanggalMulaiFormatted} sampai dengan ${tanggalSelesaiFormatted} karena ${alasan}.`;

  penutup = isEnglish
    ? "Thus, I submit this permission letter truthfully. Thank you for your attention and consideration."
    : "Demikian surat izin ini saya buat dengan sebenar-benarnya. Atas perhatian dan kebijaksanaannya, saya ucapkan terima kasih.";

  namaFile = isEnglish
    ? `Permission Letter - ${nama}.docx`
    : `Surat Izin - ${nama}.docx`;
}

if (selectedTemplateId === "surat-pengunduran-diri") {
  nama = values.nama as string;
  const jabatan = values.jabatan as string;
  const perusahaan = values.perusahaan as string;
  const tanggalTerakhir = values.tanggalTerakhir as string;
  const alasan = values.alasan as string;
  kota = values.kota as string;
  const tanggalSurat = values.tanggalSurat as string;

  tanggalSuratFormatted = isEnglish
    ? formatTanggalInggris(tanggalSurat)
    : formatTanggalIndonesia(tanggalSurat);

  const tanggalTerakhirFormatted = isEnglish
    ? formatTanggalInggris(tanggalTerakhir)
    : formatTanggalIndonesia(tanggalTerakhir);

  judul = isEnglish ? "RESIGNATION LETTER" : "SURAT PENGUNDURAN DIRI";

  tujuanSurat = perusahaan;

  identitasBaris1 = isEnglish
    ? `Name     : ${nama}`
    : `Nama     : ${nama}`;

  identitasBaris2 = isEnglish
    ? `Position : ${jabatan}`
    : `Jabatan  : ${jabatan}`;

  isiSurat = isEnglish
    ? `I hereby submit my resignation from my position as ${jabatan} at ${perusahaan}, effective ${tanggalTerakhirFormatted}.${alasan ? ` The reason for my resignation is ${alasan}.` : ""}`
    : `Dengan ini saya mengajukan pengunduran diri dari jabatan saya sebagai ${jabatan} di ${perusahaan}, terhitung efektif pada tanggal ${tanggalTerakhirFormatted}.${alasan ? ` Adapun alasan pengunduran diri saya adalah ${alasan}.` : ""}`;

  penutup = isEnglish
    ? "I would like to express my gratitude for the opportunity, trust, and experience given during my employment. I hope the company continues to grow and succeed."
    : "Saya mengucapkan terima kasih atas kesempatan, kepercayaan, dan pengalaman yang telah diberikan selama saya bekerja. Semoga perusahaan terus berkembang dan semakin sukses.";

  namaFile = isEnglish
    ? `Resignation Letter - ${nama}.docx`
    : `Surat Pengunduran Diri - ${nama}.docx`;
}

if (selectedTemplateId === "surat-lamaran-kerja") {
  nama = values.nama as string;
  const alamat = values.alamat as string;
  const email = values.email as string;
  const telepon = values.telepon as string;
  const perusahaan = values.perusahaan as string;
  const posisi = values.posisi as string;
  const pendidikan = values.pendidikan as string;
  const pengalaman = values.pengalaman as string;
  kota = values.kota as string;
  const tanggalSurat = values.tanggalSurat as string;

  tanggalSuratFormatted = isEnglish
    ? formatTanggalInggris(tanggalSurat)
    : formatTanggalIndonesia(tanggalSurat);

  judul = isEnglish ? "JOB APPLICATION LETTER" : "SURAT LAMARAN KERJA";

  tujuanSurat = isEnglish
    ? `HR Department of ${perusahaan}`
    : `Yth. HRD ${perusahaan}`;

  identitasBaris1 = isEnglish
    ? `Name    : ${nama}`
    : `Nama    : ${nama}`;

  identitasBaris2 = isEnglish
    ? `Address : ${alamat}`
    : `Alamat  : ${alamat}`;

  isiSurat = isEnglish
    ? `I hereby submit my application for the position of ${posisi} at ${perusahaan}. I have an educational background in ${pendidikan} and relevant experience in ${pengalaman}. I can be contacted by email at ${email} or by phone at ${telepon}.`
    : `Dengan ini saya mengajukan lamaran kerja untuk posisi ${posisi} di ${perusahaan}. Saya memiliki latar belakang pendidikan ${pendidikan} serta pengalaman singkat di bidang ${pengalaman}. Saya dapat dihubungi melalui email ${email} atau nomor telepon ${telepon}.`;

  penutup = isEnglish
    ? "I hope to be given the opportunity to proceed to the next stage of the recruitment process. Thank you for your attention and consideration."
    : "Besar harapan saya untuk diberikan kesempatan mengikuti tahapan seleksi berikutnya. Atas perhatian dan kesempatan yang diberikan, saya ucapkan terima kasih.";

  namaFile = isEnglish
    ? `Job Application Letter - ${nama}.docx`
    : `Surat Lamaran Kerja - ${nama}.docx`;
}

if (selectedTemplateId === "surat-kuasa") {
  const namaPemberi = values.namaPemberi as string;
  const nikPemberi = values.nikPemberi as string;
  const alamatPemberi = values.alamatPemberi as string;
  const namaPenerima = values.namaPenerima as string;
  const nikPenerima = values.nikPenerima as string;
  const alamatPenerima = values.alamatPenerima as string;
  const keperluan = values.keperluan as string;
  kota = values.kota as string;
  const tanggalSurat = values.tanggalSurat as string;

  nama = namaPemberi;

  tanggalSuratFormatted = isEnglish
    ? formatTanggalInggris(tanggalSurat)
    : formatTanggalIndonesia(tanggalSurat);

  judul = isEnglish ? "POWER OF ATTORNEY" : "SURAT KUASA";

  tujuanSurat = isEnglish
    ? "To whom it may concern"
    : "Pihak yang berkepentingan";

  identitasBaris1 = isEnglish
    ? `Grantor : ${namaPemberi}, ID Number ${nikPemberi}, Address ${alamatPemberi}`
    : `Pemberi Kuasa : ${namaPemberi}, NIK ${nikPemberi}, Alamat ${alamatPemberi}`;

  identitasBaris2 = isEnglish
    ? `Attorney : ${namaPenerima}, ID Number ${nikPenerima}, Address ${alamatPenerima}`
    : `Penerima Kuasa : ${namaPenerima}, NIK ${nikPenerima}, Alamat ${alamatPenerima}`;

  isiSurat = isEnglish
    ? `I, as the grantor, hereby authorize the attorney named above to act on my behalf for the following purpose: ${keperluan}. This power of attorney is made to be used properly and responsibly.`
    : `Saya selaku pemberi kuasa dengan ini memberikan kuasa kepada penerima kuasa tersebut di atas untuk bertindak atas nama saya dalam keperluan: ${keperluan}. Surat kuasa ini dibuat untuk dipergunakan sebagaimana mestinya dan dengan penuh tanggung jawab.`;

  penutup = isEnglish
    ? "Thus, this power of attorney is made truthfully and may be used as necessary."
    : "Demikian surat kuasa ini dibuat dengan sebenar-benarnya untuk dipergunakan sebagaimana mestinya.";

  namaFile = isEnglish
    ? `Power of Attorney - ${namaPemberi}.docx`
    : `Surat Kuasa - ${namaPemberi}.docx`;
}

if (selectedTemplateId === "surat-pernyataan") {
  nama = values.nama as string;
  const nik = values.nik as string;
  const alamat = values.alamat as string;
  const isiPernyataan = values.isiPernyataan as string;
  kota = values.kota as string;
  const tanggalSurat = values.tanggalSurat as string;

  tanggalSuratFormatted = isEnglish
    ? formatTanggalInggris(tanggalSurat)
    : formatTanggalIndonesia(tanggalSurat);

  judul = isEnglish ? "STATEMENT LETTER" : "SURAT PERNYATAAN";

  tujuanSurat = isEnglish
    ? "To whom it may concern"
    : "Pihak yang berkepentingan";

  identitasBaris1 = isEnglish
    ? `Name : ${nama}, ID Number ${nik}`
    : `Nama : ${nama}, NIK ${nik}`;

  identitasBaris2 = isEnglish
    ? `Address : ${alamat}`
    : `Alamat : ${alamat}`;

  isiSurat = isEnglish
    ? `I hereby declare that ${isiPernyataan}. This statement is made truthfully and can be used as necessary.`
    : `Dengan ini saya menyatakan bahwa ${isiPernyataan}. Pernyataan ini saya buat dengan sebenar-benarnya untuk dipergunakan sebagaimana mestinya.`;

  penutup = isEnglish
    ? "Thus, this statement letter is made truthfully and without coercion from any party."
    : "Demikian surat pernyataan ini saya buat dengan sebenar-benarnya dan tanpa paksaan dari pihak mana pun.";

  namaFile = isEnglish
    ? `Statement Letter - ${nama}.docx`
    : `Surat Pernyataan - ${nama}.docx`;
}

if (selectedTemplateId === "berita-acara") {
  const judulBeritaAcara = values.judul as string;
  const nomorBeritaAcara = (values.nomorBeritaAcara as string) || "";
  const hariTanggal = values.hariTanggal as string;
  const waktu = values.waktu as string;
  const tempat = values.tempat as string;
  const peserta = (values.peserta as string) || "";
  const uraian = values.uraian as string;
  const namaPenandatangan = values.namaPenandatangan as string;
  kota = values.kota as string;
  const tanggalSurat = values.tanggalSurat as string;

  nama = namaPenandatangan;

  tanggalSuratFormatted = isEnglish
    ? formatTanggalInggris(tanggalSurat)
    : formatTanggalIndonesia(tanggalSurat);

  judul = isEnglish ? "OFFICIAL REPORT" : "BERITA ACARA";

  namaFile = isEnglish
    ? `Official Report - ${judulBeritaAcara}.docx`
    : `Berita Acara - ${judulBeritaAcara}.docx`;
}

if (selectedTemplateId === "berita-acara") {
  const judulBeritaAcara = values.judul as string;
  const nomorBeritaAcara = (values.nomorBeritaAcara as string) || "";
  const hariTanggal = values.hariTanggal as string;
  const waktu = values.waktu as string;
  const tempat = values.tempat as string;
  const peserta = (values.peserta as string) || "";
  const uraian = values.uraian as string;
  const namaPenandatangan = values.namaPenandatangan as string;
  kota = values.kota as string;
  const tanggalSurat = values.tanggalSurat as string;

  const tanggalDokumenFormatted = isEnglish
    ? formatTanggalInggris(tanggalSurat)
    : formatTanggalIndonesia(tanggalSurat);

  const signatureImage = dataUrlToUint8Array(signatureData);

  const cell = (text: string, bold = false) =>
    new TableCell({
      margins: {
        top: 120,
        bottom: 120,
        left: 120,
        right: 120,
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text,
              bold,
              size: 24,
            }),
          ],
        }),
      ],
    });

  const row = (label: string, value: string) =>
    new TableRow({
      children: [
       new TableCell({
  width: {
    size: 30,
    type: WidthType.PERCENTAGE,
  },
  margins: {
    top: 120,
    bottom: 120,
    left: 120,
    right: 120,
  },
  children: [
    new Paragraph({
      children: [
        new TextRun({
          text: label,
          bold: true,
          size: 24,
        }),
      ],
    }),
  ],
}),

new TableCell({
  width: {
    size: 70,
    type: WidthType.PERCENTAGE,
  },
  margins: {
    top: 120,
    bottom: 120,
    left: 120,
    right: 120,
  },
  children: [
    new Paragraph({
      children: [
        new TextRun({
          text: value || "-",
          size: 24,
        }),
      ],
    }),
  ],
}),
      ],
    });

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
            children: [
              new TextRun({
                text: isEnglish ? "OFFICIAL REPORT" : "BERITA ACARA",
                bold: true,
                size: 32,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: nomorBeritaAcara ? 120 : 400 },
            children: [
              new TextRun({
                text: judulBeritaAcara.toUpperCase(),
                bold: true,
                size: 26,
              }),
            ],
          }),

          ...(nomorBeritaAcara
            ? [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 400 },
                  children: [
                    new TextRun({
                      text: `Nomor: ${nomorBeritaAcara}`,
                      size: 24,
                    }),
                  ],
                }),
              ]
            : []),

          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 300 },
            children: [
              new TextRun({
                text: isEnglish
                  ? `On ${hariTanggal}, at ${waktu}, located at ${tempat}, this official report was prepared regarding ${judulBeritaAcara}.`
                  : `Pada hari ini, ${hariTanggal}, pukul ${waktu}, bertempat di ${tempat}, telah dibuat Berita Acara mengenai ${judulBeritaAcara}.`,
                size: 24,
              }),
            ],
          }),

          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({
                text: isEnglish ? "The details are as follows:" : "Adapun rincian berita acara adalah sebagai berikut:",
                size: 24,
              }),
            ],
          }),

          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            rows: [
              row(isEnglish ? "Subject" : "Judul / Perihal", judulBeritaAcara),
              row(isEnglish ? "Day / Date" : "Hari / Tanggal", hariTanggal),
              row(isEnglish ? "Time" : "Waktu", waktu),
              row(isEnglish ? "Place" : "Tempat", tempat),
              ...(peserta
                ? [row(isEnglish ? "Related Parties / Participants" : "Pihak / Peserta yang Terlibat", peserta)]
                : []),
              row(isEnglish ? "Description" : "Uraian", uraian),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 400, after: 500 },
            children: [
              new TextRun({
                text: isEnglish
                  ? "Thus, this official report is made truthfully as documentation and may be used as necessary."
                  : "Demikian Berita Acara ini dibuat dengan sebenar-benarnya sebagai bukti/dokumentasi dan untuk dipergunakan sebagaimana mestinya.",
                size: 24,
              }),
            ],
          }),

          new Table({
            alignment: AlignmentType.RIGHT,
            width: {
              size: 3200,
              type: WidthType.DXA,
            },
            borders: {
              top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                      bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 300 },
                        children: [
                          new TextRun({
                            text: `${kota}, ${tanggalDokumenFormatted}`,
                            size: 24,
                          }),
                        ],
                      }),

                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 300 },
                        children: [
                          new TextRun({
                            text: isEnglish ? "Prepared and signed by," : "Pihak yang Menyusun Berita Acara,",
                            size: 24,
                          }),
                        ],
                      }),

                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 200 },
                        children: [
                          new ImageRun({
                            type: "png",
                            data: signatureImage,
                            transformation: {
                              width: 160,
                              height: 70,
                            },
                          }),
                        ],
                      }),

                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 100 },
                        children: [
                          new TextRun({
                            text: namaPenandatangan,
                            size: 24,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = isEnglish
    ? `Official Report - ${judulBeritaAcara}.docx`
    : `Berita Acara - ${judulBeritaAcara}.docx`;

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
  return;
}

if (selectedTemplateId === "notulen-rapat") {
  const judulRapat = values.judulRapat as string;
  const tanggalRapat = values.tanggalRapat as string;
  const waktuRapat = values.waktuRapat as string;
  const tempatRapat = values.tempatRapat as string;
  const pimpinanRapat = values.pimpinanRapat as string;
  const notulis = values.notulis as string;
  const pesertaRapat = values.pesertaRapat as string;
  const agenda = values.agenda as string;
  const kota = values.kota as string;
  const tanggalDokumen = values.tanggalDokumen as string;

  const pembahasanRows = notulenEntries
    .map((entry) => ({
      pembahasan: entry.pembahasan.trim(),
      keputusan: entry.keputusan.trim(),
      tindakLanjut: entry.tindakLanjut.trim(),
    }))
    .filter(
      (entry) =>
        entry.pembahasan !== "" ||
        entry.keputusan !== "" ||
        entry.tindakLanjut !== ""
    );

  if (pembahasanRows.length === 0) {
    alert(t.discussionEntryRequired);
    return;
  }

  if (pembahasanRows.some((entry) => entry.pembahasan === "")) {
    alert(t.discussionColumnRequired);
    return;
  }

  const tanggalRapatFormatted = isEnglish
    ? formatTanggalInggris(tanggalRapat)
    : formatTanggalIndonesia(tanggalRapat);

  const tanggalDokumenFormatted = isEnglish
    ? formatTanggalInggris(tanggalDokumen)
    : formatTanggalIndonesia(tanggalDokumen);

  const signatureImage = dataUrlToUint8Array(signatureData);

  const borderNone = {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  };

  const textParagraphs = (text: string, bold = false, size = 22) =>
    (text || "-").split("\n").map(
      (line) =>
        new Paragraph({
          spacing: { after: 80 },
          children: [
            new TextRun({
              text: line || " ",
              bold,
              size,
            }),
          ],
        })
    );

  const metadataCell = (text: string, bold = false, width = 50) =>
    new TableCell({
      width: {
        size: width,
        type: WidthType.PERCENTAGE,
      },
      margins: {
        top: 120,
        bottom: 120,
        left: 120,
        right: 120,
      },
      children: textParagraphs(text, bold, 24),
    });

  const metadataRow = (label: string, value: string) =>
    new TableRow({
      children: [
        metadataCell(label, true, 30),
        metadataCell(value || "-", false, 70),
      ],
    });

  const discussionCell = (
    text: string,
    bold = false,
    width = 25,
    alignment: typeof AlignmentType.CENTER | typeof AlignmentType.LEFT =
      AlignmentType.LEFT
  ) =>
    new TableCell({
      width: {
        size: width,
        type: WidthType.PERCENTAGE,
      },
      margins: {
        top: 120,
        bottom: 120,
        left: 120,
        right: 120,
      },
      children: [
        new Paragraph({
          alignment,
          children: [
            new TextRun({
              text: text || "-",
              bold,
              size: 22,
            }),
          ],
        }),
      ],
    });

  const discussionTextCell = (text: string, width: number) =>
    new TableCell({
      width: {
        size: width,
        type: WidthType.PERCENTAGE,
      },
      margins: {
        top: 120,
        bottom: 120,
        left: 120,
        right: 120,
      },
      children: textParagraphs(text || "-", false, 22),
    });

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
            children: [
              new TextRun({
                text: isEnglish ? "MEETING MINUTES" : "NOTULEN RAPAT",
                bold: true,
                size: 32,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
            children: [
              new TextRun({
                text: judulRapat.toUpperCase(),
                bold: true,
                size: 26,
              }),
            ],
          }),

          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            layout: TableLayoutType.FIXED,
            rows: [
              metadataRow(
                isEnglish ? "Meeting Title" : "Judul / Nama Rapat",
                judulRapat
              ),
              metadataRow(
                isEnglish ? "Date" : "Tanggal Rapat",
                tanggalRapatFormatted
              ),
              metadataRow(isEnglish ? "Time" : "Waktu Rapat", waktuRapat),
              metadataRow(
                isEnglish ? "Place / Media" : "Tempat / Media Rapat",
                tempatRapat
              ),
              metadataRow(
                isEnglish ? "Meeting Chair" : "Pimpinan Rapat",
                pimpinanRapat
              ),
              metadataRow(isEnglish ? "Minutes Prepared By" : "Notulis", notulis),
              metadataRow(
                isEnglish ? "Participants" : "Peserta Rapat",
                pesertaRapat
              ),
              metadataRow(isEnglish ? "Agenda" : "Agenda Rapat", agenda),
            ],
          }),

          new Paragraph({
            spacing: { before: 400, after: 160 },
            children: [
              new TextRun({
                text: isEnglish
                  ? "Discussion, Decisions, and Follow-up Actions"
                  : "Pembahasan, Keputusan, dan Tindak Lanjut",
                bold: true,
                size: 26,
              }),
            ],
          }),

          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            layout: TableLayoutType.FIXED,
            rows: [
              new TableRow({
                children: [
                  discussionCell("No.", true, 8, AlignmentType.CENTER),
                  discussionCell(
                    isEnglish ? "Discussion" : "Pembahasan",
                    true,
                    37,
                    AlignmentType.CENTER
                  ),
                  discussionCell(
                    isEnglish ? "Decision" : "Keputusan",
                    true,
                    27,
                    AlignmentType.CENTER
                  ),
                  discussionCell(
                    isEnglish ? "Follow-up" : "Tindak Lanjut",
                    true,
                    28,
                    AlignmentType.CENTER
                  ),
                ],
              }),

              ...pembahasanRows.map(
                (entry, index) =>
                  new TableRow({
                    children: [
                      discussionCell(
                        String(index + 1),
                        false,
                        8,
                        AlignmentType.CENTER
                      ),
                      discussionTextCell(entry.pembahasan, 37),
                      discussionTextCell(entry.keputusan || "-", 27),
                      discussionTextCell(entry.tindakLanjut || "-", 28),
                    ],
                  })
              ),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 400, after: 500 },
            children: [
              new TextRun({
                text: isEnglish
                  ? "These meeting minutes are prepared as official documentation of the discussion, decisions, and follow-up actions agreed upon in the meeting."
                  : "Demikian notulen rapat ini disusun sebagai dokumentasi resmi atas pembahasan, keputusan, dan tindak lanjut yang disepakati dalam rapat.",
                size: 24,
              }),
            ],
          }),

          new Table({
            alignment: AlignmentType.RIGHT,
            width: {
              size: 3200,
              type: WidthType.DXA,
            },
            borders: {
              top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              insideHorizontal: {
                style: BorderStyle.NONE,
                size: 0,
                color: "FFFFFF",
              },
              insideVertical: {
                style: BorderStyle.NONE,
                size: 0,
                color: "FFFFFF",
              },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    borders: borderNone,
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 300 },
                        children: [
                          new TextRun({
                            text: `${kota}, ${tanggalDokumenFormatted}`,
                            size: 24,
                          }),
                        ],
                      }),

                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 300 },
                        children: [
                          new TextRun({
                            text: isEnglish ? "Minutes prepared by," : "Notulis,",
                            size: 24,
                          }),
                        ],
                      }),

                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 200 },
                        children: [
                          new ImageRun({
                            type: "png",
                            data: signatureImage,
                            transformation: {
                              width: 160,
                              height: 70,
                            },
                          }),
                        ],
                      }),

                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 100 },
                        children: [
                          new TextRun({
                            text: notulis,
                            size: 24,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = isEnglish
    ? `Meeting Minutes - ${judulRapat}.docx`
    : `Notulen Rapat - ${judulRapat}.docx`;

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
  return;
}


if (selectedTemplateId === "proposal-kegiatan") {
  const namaKegiatan = values.namaKegiatan as string;
  const latarBelakang = values.latarBelakang as string;
  const tujuan = values.tujuan as string;
  const waktuKegiatan = values.waktuKegiatan as string;
  const tempatKegiatan = values.tempatKegiatan as string;
  const sasaranPeserta = values.sasaranPeserta as string;
  const anggaran = (values.anggaran as string) || "";
  const namaPenanggungJawab = values.namaPenanggungJawab as string;
  const kota = values.kota as string;
  const tanggalDokumen = values.tanggalDokumen as string;

  const tanggalDokumenFormatted = isEnglish
    ? formatTanggalInggris(tanggalDokumen)
    : formatTanggalIndonesia(tanggalDokumen);

  const signatureImage = dataUrlToUint8Array(signatureData);

  const paragraphFromText = (text: string) =>
    (text || "-").split("\n").map(
      (line) =>
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 180 },
          children: [
            new TextRun({
              text: line || " ",
              size: 24,
            }),
          ],
        })
    );

  const sectionTitle = (text: string) =>
    new Paragraph({
      spacing: { before: 360, after: 160 },
      children: [
        new TextRun({
          text,
          bold: true,
          size: 26,
        }),
      ],
    });

  const infoRow = (label: string, value: string) =>
    new TableRow({
      children: [
        new TableCell({
          width: {
            size: 30,
            type: WidthType.PERCENTAGE,
          },
          margins: {
            top: 120,
            bottom: 120,
            left: 120,
            right: 120,
          },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: label,
                  bold: true,
                  size: 24,
                }),
              ],
            }),
          ],
        }),

        new TableCell({
          width: {
            size: 70,
            type: WidthType.PERCENTAGE,
          },
          margins: {
            top: 120,
            bottom: 120,
            left: 120,
            right: 120,
          },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: value || "-",
                  size: 24,
                }),
              ],
            }),
          ],
        }),
      ],
    });

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
            children: [
              new TextRun({
                text: isEnglish ? "ACTIVITY PROPOSAL" : "PROPOSAL KEGIATAN",
                bold: true,
                size: 32,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 500 },
            children: [
              new TextRun({
                text: namaKegiatan.toUpperCase(),
                bold: true,
                size: 28,
              }),
            ],
          }),

          sectionTitle(isEnglish ? "A. Background" : "A. Latar Belakang"),
          ...paragraphFromText(latarBelakang),

          sectionTitle(
            isEnglish ? "B. Purpose and Objectives" : "B. Maksud dan Tujuan"
          ),
          ...paragraphFromText(tujuan),

          sectionTitle(
            isEnglish
              ? "C. Time and Place of Activity"
              : "C. Waktu dan Tempat Kegiatan"
          ),

          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            layout: TableLayoutType.FIXED,
            rows: [
              infoRow(isEnglish ? "Time" : "Waktu", waktuKegiatan),
              infoRow(isEnglish ? "Place" : "Tempat", tempatKegiatan),
            ],
          }),

          sectionTitle(
            isEnglish ? "D. Target Participants" : "D. Sasaran Peserta"
          ),
          ...paragraphFromText(sasaranPeserta),

          sectionTitle(
            isEnglish ? "E. Budget Plan" : "E. Rencana Anggaran"
          ),
          ...paragraphFromText(anggaran || "-"),

          sectionTitle(isEnglish ? "F. Closing" : "F. Penutup"),

          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 500 },
            children: [
              new TextRun({
                text: isEnglish
                  ? "Thus, this activity proposal is prepared to serve as a reference for the implementation of the activity. Thank you for your attention and support."
                  : "Demikian proposal kegiatan ini disusun sebagai acuan pelaksanaan kegiatan. Atas perhatian dan dukungan yang diberikan, kami ucapkan terima kasih.",
                size: 24,
              }),
            ],
          }),

          new Table({
            alignment: AlignmentType.RIGHT,
            width: {
              size: 3200,
              type: WidthType.DXA,
            },
            borders: {
              top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              insideHorizontal: {
                style: BorderStyle.NONE,
                size: 0,
                color: "FFFFFF",
              },
              insideVertical: {
                style: BorderStyle.NONE,
                size: 0,
                color: "FFFFFF",
              },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                      bottom: {
                        style: BorderStyle.NONE,
                        size: 0,
                        color: "FFFFFF",
                      },
                      left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 300 },
                        children: [
                          new TextRun({
                            text: `${kota}, ${tanggalDokumenFormatted}`,
                            size: 24,
                          }),
                        ],
                      }),

                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 300 },
                        children: [
                          new TextRun({
                            text: isEnglish
                              ? "Person in charge,"
                              : "Penanggung Jawab,",
                            size: 24,
                          }),
                        ],
                      }),

                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 200 },
                        children: [
                          new ImageRun({
                            type: "png",
                            data: signatureImage,
                            transformation: {
                              width: 160,
                              height: 70,
                            },
                          }),
                        ],
                      }),

                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 100 },
                        children: [
                          new TextRun({
                            text: namaPenanggungJawab,
                            size: 24,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = isEnglish
    ? `Activity Proposal - ${namaKegiatan}.docx`
    : `Proposal Kegiatan - ${namaKegiatan}.docx`;

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
  return;
}


if (selectedTemplateId === "invoice") {
  const namaBisnis = values.namaBisnis as string;
const namaPelanggan = values.namaPelanggan as string;
const nomorInvoice = values.nomorInvoice as string;
const tanggalInvoice = values.tanggalInvoice as string;
const mataUang = (values.mataUang as string) || "IDR";
const catatan = (values.catatan as string) || "";

  const invoiceRows = invoiceItems
    .map((item) => {
      const jumlah = Number(item.jumlah || 0);
      const hargaSatuan = Number(item.hargaSatuan || 0);

      return {
        deskripsi: item.deskripsi.trim(),
        jumlah,
        hargaSatuan,
        total: jumlah * hargaSatuan,
      };
    })
    .filter(
      (item) =>
        item.deskripsi !== "" || item.jumlah > 0 || item.hargaSatuan > 0
    );

  if (invoiceRows.length === 0) {
    alert(t.invoiceItemRequired);
    return;
  }

  if (invoiceRows.some((item) => item.deskripsi === "")) {
    alert(t.invoiceDescriptionRequired);
    return;
  }

  if (
    invoiceRows.some(
      (item) =>
        !Number.isFinite(item.jumlah) ||
        !Number.isFinite(item.hargaSatuan) ||
        item.jumlah <= 0 ||
        item.hargaSatuan <= 0
    )
  ) {
    alert(t.invoiceAmountRequired);
    return;
  }

  const totalTagihan = invoiceRows.reduce(
    (sum, item) => sum + item.total,
    0
  );

  const tanggalInvoiceFormatted = isEnglish
    ? formatTanggalInggris(tanggalInvoice)
    : formatTanggalIndonesia(tanggalInvoice);

  const formatUang = (angka: number) =>
  formatCurrencyAmount(angka, mataUang, isEnglish);

  const formatAngka = (angka: number) =>
    new Intl.NumberFormat(isEnglish ? "en-US" : "id-ID", {
      maximumFractionDigits: 2,
    }).format(angka);

  const signatureImage = dataUrlToUint8Array(signatureData);

  const borderNone = {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  };

  const paragraphFromText = (text: string, size = 24, bold = false) =>
    (text || "-").split("\n").map(
      (line) =>
        new Paragraph({
          spacing: { after: 80 },
          children: [
            new TextRun({
              text: line || " ",
              bold,
              size,
            }),
          ],
        })
    );

  const infoCell = (text: string, bold = false, width = 50) =>
    new TableCell({
      width: {
        size: width,
        type: WidthType.PERCENTAGE,
      },
      margins: {
        top: 120,
        bottom: 120,
        left: 120,
        right: 120,
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: text || "-",
              bold,
              size: 24,
            }),
          ],
        }),
      ],
    });

  const itemCell = (text: string, bold = false, width = 25) =>
    new TableCell({
      width: {
        size: width,
        type: WidthType.PERCENTAGE,
      },
      margins: {
        top: 120,
        bottom: 120,
        left: 120,
        right: 120,
      },
      children: paragraphFromText(text || "-", 22, bold),
    });

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
            children: [
              new TextRun({
                text: "INVOICE",
                bold: true,
                size: 36,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 500 },
            children: [
              new TextRun({
                text: nomorInvoice,
                size: 24,
              }),
            ],
          }),

          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            layout: TableLayoutType.FIXED,
            rows: [
              new TableRow({
                children: [
                  infoCell(isEnglish ? "Seller / Business" : "Penjual / Bisnis", true, 30),
                  infoCell(namaBisnis, false, 70),
                ],
              }),
              new TableRow({
                children: [
                  infoCell(isEnglish ? "Customer" : "Pelanggan", true, 30),
                  infoCell(namaPelanggan, false, 70),
                ],
              }),
              new TableRow({
                children: [
                  infoCell(isEnglish ? "Invoice Date" : "Tanggal Invoice", true, 30),
                  infoCell(tanggalInvoiceFormatted, false, 70),
                ],
              }),
new TableRow({
  children: [
    infoCell(isEnglish ? "Currency" : "Mata Uang", true, 30),
    infoCell(getCurrencyLabel(mataUang), false, 70),
  ],
}),

            ],
          }),

          new Paragraph({
            spacing: { before: 500, after: 160 },
            children: [
              new TextRun({
                text: isEnglish ? "Billing Details" : "Rincian Tagihan",
                bold: true,
                size: 26,
              }),
            ],
          }),

          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            layout: TableLayoutType.FIXED,
            rows: [
              new TableRow({
                children: [
                  itemCell(isEnglish ? "Description" : "Deskripsi", true, 40),
                  itemCell(isEnglish ? "Qty" : "Jumlah", true, 15),
                  itemCell(isEnglish ? "Unit Price" : "Harga Satuan", true, 22),
                  itemCell("Total", true, 23),
                ],
              }),
              ...invoiceRows.map(
                (item) =>
                  new TableRow({
                    children: [
                      itemCell(item.deskripsi, false, 40),
                      itemCell(formatAngka(item.jumlah), false, 15),
                      itemCell(formatUang(item.hargaSatuan), false, 22),
                      itemCell(formatUang(item.total), false, 23),
                    ],
                  })
              ),
              new TableRow({
                children: [
                  itemCell(" ", false, 40),
                  itemCell(" ", false, 15),
                  itemCell(isEnglish ? "Grand Total" : "Total Tagihan", true, 22),
                  itemCell(formatUang(totalTagihan), true, 23),
                ],
              }),
            ],
          }),

          ...(catatan
            ? [
                new Paragraph({
                  spacing: { before: 400, after: 120 },
                  children: [
                    new TextRun({
                      text: isEnglish ? "Payment Notes" : "Catatan Pembayaran",
                      bold: true,
                      size: 26,
                    }),
                  ],
                }),
                ...paragraphFromText(catatan),
              ]
            : []),

          new Paragraph({
            spacing: { before: 500, after: 500 },
            children: [
              new TextRun({
                text: isEnglish
                  ? "Thank you for your trust and cooperation."
                  : "Terima kasih atas kepercayaan dan kerja samanya.",
                size: 24,
              }),
            ],
          }),

          new Table({
            alignment: AlignmentType.RIGHT,
            width: {
              size: 3200,
              type: WidthType.DXA,
            },
            borders: {
              top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    borders: borderNone,
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 300 },
                        children: [
                          new TextRun({
                            text: isEnglish ? "Seller," : "Penjual,",
                            size: 24,
                          }),
                        ],
                      }),

                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 200 },
                        children: [
                          new ImageRun({
                            type: "png",
                            data: signatureImage,
                            transformation: {
                              width: 160,
                              height: 70,
                            },
                          }),
                        ],
                      }),

                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 100 },
                        children: [
                          new TextRun({
                            text: namaBisnis,
                            size: 24,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = `Invoice - ${nomorInvoice}.docx`;

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
  return;
}

if (selectedTemplateId === "cv-sederhana") {
  const nama = values.nama as string;
  const profesi = values.profesi as string;
  const email = values.email as string;
  const telepon = values.telepon as string;
  const ringkasan = values.ringkasan as string;
  const pendidikan = values.pendidikan as string;
  const pengalaman = (values.pengalaman as string) || "";
  const keahlian = values.keahlian as string;

  const signatureImage = dataUrlToUint8Array(signatureData);

const borderNone = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

  const sectionTitle = (text: string) =>
    new Paragraph({
      spacing: { before: 360, after: 160 },
      children: [
        new TextRun({
          text,
          bold: true,
          size: 26,
        }),
      ],
    });

  const paragraphFromText = (text: string) =>
    (text || "-").split("\n").map(
      (line) =>
        new Paragraph({
          spacing: { after: 120 },
          children: [
            new TextRun({
              text: line || " ",
              size: 24,
            }),
          ],
        })
    );

  const contactCell = (label: string, value: string) =>
    new TableCell({
      width: {
        size: 50,
        type: WidthType.PERCENTAGE,
      },
      margins: {
        top: 120,
        bottom: 120,
        left: 120,
        right: 120,
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: `${label}: `,
              bold: true,
              size: 24,
            }),
            new TextRun({
              text: value || "-",
              size: 24,
            }),
          ],
        }),
      ],
    });

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 },
            children: [
              new TextRun({
                text: nama.toUpperCase(),
                bold: true,
                size: 36,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 },
            children: [
              new TextRun({
                text: profesi,
                size: 26,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 500 },
            children: [
              new TextRun({
  text: "CURRICULUM VITAE",
  bold: true,
  size: 28,
}),
            ],
          }),

          new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE,
            },
            layout: TableLayoutType.FIXED,
            rows: [
              new TableRow({
                children: [
                  contactCell("Email", email),
                  contactCell(isEnglish ? "Phone" : "Nomor HP", telepon),
                ],
              }),
            ],
          }),

          sectionTitle(isEnglish ? "Profile Summary" : "Ringkasan Profil"),
          ...paragraphFromText(ringkasan),

          sectionTitle(isEnglish ? "Education" : "Riwayat Pendidikan"),
          ...paragraphFromText(pendidikan),

          sectionTitle(isEnglish ? "Work Experience" : "Pengalaman Kerja"),
          ...paragraphFromText(pengalaman || "-"),

          sectionTitle(isEnglish ? "Skills" : "Keahlian"),
          ...paragraphFromText(keahlian),

          new Table({
  alignment: AlignmentType.RIGHT,
  width: {
    size: 3200,
    type: WidthType.DXA,
  },
  borders: {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  },
  rows: [
    new TableRow({
      children: [
        new TableCell({
          borders: borderNone,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 500, after: 200 },
              children: [
                new TextRun({
                  text: isEnglish ? "Prepared by," : "Dibuat oleh,",
                  size: 24,
                }),
              ],
            }),

            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 100, after: 100 },
              children: [
                new ImageRun({
                  type: "png",
                  data: signatureImage,
                  transformation: {
                    width: 160,
                    height: 70,
                  },
                }),
              ],
            }),

            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: nama,
                  size: 24,
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ],
}),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = `Curriculum Vitae - ${nama}.docx`;

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
  return;
}

const signatureImage = dataUrlToUint8Array(signatureData);

  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
            children: [
              new TextRun({
                text: judul,
                bold: true,
                size: 28,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { after: 400 },
            children: [
              new TextRun({
                text: `${kota}, ${tanggalSuratFormatted}`,
                size: 24,
              }),
            ],
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: isEnglish ? "To:" : "Kepada Yth.",
                size: 24,
              }),
            ],
          }),

          new Paragraph({
            spacing: { after: 300 },
            children: [
              new TextRun({
                text: tujuanSurat,
                size: 24,
              }),
            ],
          }),

          new Paragraph({
            spacing: { after: 300 },
            children: [
              new TextRun({
                text: isEnglish ? "Respectfully," : "Dengan hormat,",
                size: 24,
              }),
            ],
          }),

          new Paragraph({
            spacing: { after: 300 },
            children: [
              new TextRun({
                text: isEnglish
                  ? "I, the undersigned below:"
                  : "Saya yang bertanda tangan di bawah ini:",
                size: 24,
              }),
            ],
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: identitasBaris1,
                size: 24,
              }),
            ],
          }),

          new Paragraph({
            spacing: { after: 300 },
            children: [
              new TextRun({
                text: identitasBaris2,
                size: 24,
              }),
            ],
          }),

          new Paragraph({
            spacing: { after: 300 },
            children: [
              new TextRun({
                text: isiSurat,
                size: 24,
              }),
            ],
          }),

          new Paragraph({
            spacing: { after: 500 },
            children: [
              new TextRun({
                text: penutup,
                size: 24,
              }),
            ],
          }),

          new Table({
  alignment: AlignmentType.RIGHT,
  width: {
    size: 3200,
    type: WidthType.DXA,
  },
  borders: {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  },
  rows: [
    new TableRow({
      children: [
        new TableCell({
          borders: {
            top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 600 },
              children: [
                new TextRun({
                  text: isEnglish ? "Sincerely," : "Hormat saya,",
                  size: 24,
                }),
              ],
            }),

            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 200 },
              children: [
                new ImageRun({
                  type: "png",
                  data: signatureImage,
                  transformation: {
                    width: 160,
                    height: 70,
                  },
                }),
              ],
            }),

            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 100 },
              children: [
                new TextRun({
                  text: nama,
                  size: 24,
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ],
}),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = namaFile;

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
}

  return (
    <main className="min-h-screen bg-slate-50 text-slate-900">
      <section className="mx-auto max-w-6xl px-6 py-10">
        <div className="mb-8 rounded-3xl bg-white p-8 shadow-sm">
  <div className="mb-3 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
    <p className="text-sm font-semibold uppercase tracking-wide text-blue-600">
      DokumenJadi
    </p>

    <div className="inline-flex w-fit rounded-xl border border-slate-200 bg-slate-50 p-1">
      <button
        type="button"
        onClick={() => setUiLanguage("id")}
        className={`rounded-lg px-4 py-2 text-sm font-semibold transition ${
          uiLanguage === "id"
            ? "bg-blue-600 text-white"
            : "text-slate-600 hover:bg-white"
        }`}
      >
        Indonesia
      </button>

      <button
        type="button"
        onClick={() => setUiLanguage("en")}
        className={`rounded-lg px-4 py-2 text-sm font-semibold transition ${
          uiLanguage === "en"
            ? "bg-blue-600 text-white"
            : "text-slate-600 hover:bg-white"
        }`}
      >
        English
      </button>
    </div>
  </div>

  <h1 className="mb-4 max-w-3xl text-4xl font-bold tracking-tight md:text-5xl">
    {t.heroTitle}
  </h1>

  <p className="max-w-2xl text-lg text-slate-600">
    {t.heroDescription}
  </p>
</div>

        <div className="grid gap-6 lg:grid-cols-[1fr_1.2fr]">
          <div>
            <h2 className="mb-4 text-2xl font-bold">{t.chooseTemplate}</h2>

            <div className="grid gap-3">
              {templates.map((template) => {
                const isSelected = template.id === selectedTemplateId;

                return (
                  <button
                    key={template.id}
                    onClick={() => setSelectedTemplateId(template.id)}
                    className={`rounded-2xl border p-5 text-left shadow-sm transition ${
                      isSelected
                        ? "border-blue-500 bg-blue-50"
                        : "border-slate-200 bg-white hover:border-blue-300"
                    }`}
                  >
                    <h3 className="mb-1 text-lg font-semibold">
                      {getTemplateName(template)}
                    </h3>
                    <p className="text-sm leading-6 text-slate-600">
                      {getTemplateDescription(template)}
                    </p>
                  </button>
                );
              })}
            </div>
          </div>

          <div className="rounded-3xl bg-white p-6 shadow-sm">
            <h2 className="mb-2 text-2xl font-bold">
             {selectedTemplate ? getTemplateName(selectedTemplate) : ""}
            </h2>

            <p className="mb-6 text-slate-600">
              {t.formHelper}
            </p>

            <form ref={formRef} onSubmit={handleSubmit} className="grid gap-4">
              {selectedTemplate?.fields.map((field) => (
                <div key={field.name}>
                  <label className="mb-2 block text-sm font-medium text-slate-700">
                    {getFieldLabel(field)}
                    {field.required && <span className="text-red-500"> *</span>}
                  </label>

                  {field.type === "textarea" ? (
                    <textarea
                      name={field.name}
                      required={field.required}
                      rows={4}
                      className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none focus:border-blue-500"
                    />
                  ) : (
                    <input
                      type={field.type}
                      name={field.name}
                      required={field.required}
                      className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none focus:border-blue-500"
                    />
                  )}
                </div>
              ))}

{selectedTemplateId === "notulen-rapat" && (
  <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
    <div className="mb-4 flex items-center justify-between gap-3">
      <div>
        <h3 className="text-lg font-semibold text-slate-900">
          {t.discussionListTitle}
        </h3>
        <p className="text-sm text-slate-500">
          {t.discussionListDescription}
        </p>
      </div>

      <button
        type="button"
        onClick={addNotulenEntry}
        className="rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-700"
      >
        {t.addEntry}
      </button>
    </div>

    <div className="grid gap-4">
      {notulenEntries.map((entry, index) => (
        <div
          key={index}
          className="rounded-2xl border border-slate-200 bg-white p-4"
        >
          <div className="mb-3 flex items-center justify-between">
            <h4 className="font-semibold text-slate-800">
              {t.entry} {index + 1}
            </h4>

            {notulenEntries.length > 1 && (
              <button
                type="button"
                onClick={() => removeNotulenEntry(index)}
                className="rounded-lg border border-red-200 px-3 py-1 text-sm font-semibold text-red-600 hover:bg-red-50"
              >
                {t.remove}
              </button>
            )}
          </div>

          <div className="grid gap-3">
            <div>
              <label className="mb-2 block text-sm font-medium text-slate-700">
    {t.discussion} <span className="text-red-500">*</span>
  </label>
              <textarea
                value={entry.pembahasan}
                onChange={(event) =>
                  updateNotulenEntry(index, "pembahasan", event.target.value)
                }
                rows={3}
                className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none focus:border-blue-500"
              />
            </div>

            <div>
              <label className="mb-2 block text-sm font-medium text-slate-700">
    {t.decision}
  </label>
              <textarea
                value={entry.keputusan}
                onChange={(event) =>
                  updateNotulenEntry(index, "keputusan", event.target.value)
                }
                rows={3}
                className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none focus:border-blue-500"
              />
            </div>

            <div>
              <label className="mb-2 block text-sm font-medium text-slate-700">
    {t.followUp}
  </label>
              <textarea
                value={entry.tindakLanjut}
                onChange={(event) =>
                  updateNotulenEntry(index, "tindakLanjut", event.target.value)
                }
                rows={3}
                className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none focus:border-blue-500"
              />
            </div>
          </div>
        </div>
      ))}
    </div>
  </div>
)}

{selectedTemplateId === "invoice" && (
  <div>
    <label className="mb-2 block text-sm font-medium text-slate-700">
    {t.currency} <span className="text-red-500">*</span>
  </label>

    <select
      name="mataUang"
      defaultValue="IDR"
      required
      className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none focus:border-blue-500"
    >
      {currencyOptions.map((currency) => (
        <option key={currency.value} value={currency.value}>
          {currency.label}
        </option>
      ))}
    </select>
  </div>
)}

{selectedTemplateId === "invoice" && (
  <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
    <div className="mb-4 flex items-center justify-between gap-3">
      <div>
        <h3 className="text-lg font-semibold text-slate-900">
          {t.invoiceItemListTitle}
        </h3>
        <p className="text-sm text-slate-500">
          {t.invoiceItemListDescription}
        </p>
      </div>

      <button
        type="button"
        onClick={addInvoiceItem}
        className="rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-700"
      >
        {t.addItem}
      </button>
    </div>

    <div className="grid gap-4">
      {invoiceItems.map((item, index) => (
        <div
          key={index}
          className="rounded-2xl border border-slate-200 bg-white p-4"
        >
          <div className="mb-3 flex items-center justify-between">
            <h4 className="font-semibold text-slate-800">
              {t.item} {index + 1}
            </h4>

            {invoiceItems.length > 1 && (
              <button
                type="button"
                onClick={() => removeInvoiceItem(index)}
                className="rounded-lg border border-red-200 px-3 py-1 text-sm font-semibold text-red-600 hover:bg-red-50"
              >
                {t.remove}
              </button>
            )}
          </div>

          <div className="grid gap-3">
            <div>
              <label className="mb-2 block text-sm font-medium text-slate-700">
    {t.productDescription} <span className="text-red-500">*</span>
  </label>
              <textarea
                value={item.deskripsi}
                onChange={(event) =>
                  updateInvoiceItem(index, "deskripsi", event.target.value)
                }
                rows={3}
                required
                className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none focus:border-blue-500"
              />
            </div>

            <div className="grid gap-3 sm:grid-cols-2">
              <div>
                <label className="mb-2 block text-sm font-medium text-slate-700">
    {t.quantity} <span className="text-red-500">*</span>
  </label>
                <input
                  type="number"
                  value={item.jumlah}
                  onChange={(event) =>
                    updateInvoiceItem(index, "jumlah", event.target.value)
                  }
                  min="0"
                  step="any"
                  required
                  className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none focus:border-blue-500"
                />
              </div>

              <div>
                <label className="mb-2 block text-sm font-medium text-slate-700">
    {t.unitPrice} <span className="text-red-500">*</span>
  </label>
                <input
                  type="number"
                  value={item.hargaSatuan}
                  onChange={(event) =>
                    updateInvoiceItem(index, "hargaSatuan", event.target.value)
                  }
                  min="0"
                  step="any"
                  required
                  className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none focus:border-blue-500"
                />
              </div>
            </div>
          </div>
        </div>
      ))}
    </div>
  </div>
)}

              <div>
                <label className="mb-2 block text-sm font-medium text-slate-700">
    {t.documentLanguage} <span className="text-red-500">*</span>
  </label>

                <select
                  name="bahasaDokumen"
                  defaultValue="id"
                  required
                  className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none focus:border-blue-500"
                >
                  <option value="id">{t.indonesian}</option>
                  <option value="en">{t.english}</option>
                </select>
              </div>

<div>
  <label className="mb-2 block text-sm font-medium text-slate-700">
    {t.signature} <span className="text-red-500">*</span>
  </label>

  <div className="rounded-xl border border-slate-300 bg-white p-3">
    <canvas
      ref={signatureCanvasRef}
      width={700}
      height={220}
      onPointerDown={startSignature}
      onPointerMove={drawSignature}
      onPointerUp={endSignature}
      onPointerLeave={endSignature}
      className="h-40 w-full touch-none cursor-crosshair rounded-lg bg-white"
    />
  </div>

  <div className="mt-2 flex items-center justify-between gap-3">
    <p className="text-sm text-slate-500">
  {t.signatureInstruction}
</p>

    <button
      type="button"
      onClick={clearSignature}
      className="rounded-lg border border-slate-300 px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50"
    >
      {t.clearSignature}
    </button>
  </div>
</div>

              <div className="mt-4 grid gap-3 sm:grid-cols-2">
  <button
    type="button"
    onClick={() => setShowClearConfirm(true)}
    className="rounded-xl border border-slate-300 px-5 py-3 font-semibold text-slate-700 transition hover:bg-slate-50"
  >
    {t.clearInputs}
  </button>

  <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
    <button
      type="submit"
      name="downloadType"
      value="word"
      className="rounded-xl bg-blue-600 px-5 py-3 font-semibold text-white transition hover:bg-blue-700"
    >
      {t.downloadWord}
    </button>

    <button
      type="submit"
      name="downloadType"
      value="pdf"
      className="rounded-xl bg-red-600 px-5 py-3 font-semibold text-white transition hover:bg-red-700"
    >
      {t.downloadPdf}
    </button>
  </div>
</div>
</form>
            {showClearConfirm && (
  <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4">
    <div className="w-full max-w-md rounded-2xl bg-white p-6 shadow-xl">
      <h3 className="mb-2 text-xl font-bold text-slate-900">
        {t.clearModalTitle}
      </h3>

      <p className="mb-6 text-sm leading-6 text-slate-600">
        {t.clearModalBody}
      </p>

      <div className="grid gap-3 sm:grid-cols-2">
        <button
          type="button"
          onClick={() => setShowClearConfirm(false)}
          className="rounded-xl border border-slate-300 px-4 py-3 font-semibold text-slate-700 hover:bg-slate-50"
        >
          {t.no}
        </button>

        <button
          type="button"
          onClick={clearAllInputs}
          className="rounded-xl bg-red-600 px-4 py-3 font-semibold text-white hover:bg-red-700"
        >
          {t.yesClear}
        </button>
      </div>
    </div>
  </div>
)}
          </div>
        </div>
      </section>
    </main>
  );
}