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

function handleDownloadPdf(
  values: Record<string, FormDataEntryValue>,
  isEnglish: boolean
) {
  const pdf = new jsPDF("p", "mm", "a4");

  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();

  const marginLeft = 25;
  const marginRight = 25;
  const contentWidth = pageWidth - marginLeft - marginRight;

  let y = 28;

  const getValue = (key: string) => String(values[key] || "").trim();

  const safeFileName = (text: string) =>
    text.replace(/[\\/:*?"<>|]/g, "-").trim() || "DokumenJadi";

  function checkPage(requiredSpace = 12) {
    if (y + requiredSpace > pageHeight - 22) {
      pdf.addPage();
      y = 28;
    }
  }

  function addCenteredText(text: string, size = 14, bold = false, after = 7) {
    checkPage(after + 5);

    pdf.setFont("times", bold ? "bold" : "normal");
    pdf.setFontSize(size);
    pdf.text(text, pageWidth / 2, y, { align: "center" });

    y += after;
  }

  function addSectionTitle(text: string) {
    checkPage(12);

    y += 5;
    pdf.setFont("times", "bold");
    pdf.setFontSize(13);
    pdf.text(text, marginLeft, y);

    y += 8;
  }

  function addParagraph(text: string) {
    pdf.setFont("times", "normal");
    pdf.setFontSize(12);

    const paragraphs = (text || "-").split("\n");

    paragraphs.forEach((paragraph) => {
      const lines = pdf.splitTextToSize(
        paragraph || " ",
        contentWidth
      ) as string[];

      lines.forEach((line) => {
        checkPage(8);
        pdf.text(line, marginLeft, y);
        y += 7;
      });

      y += 2;
    });
  }

  function addInfoTable(rows: { label: string; value: string }[]) {
    const labelWidth = 48;
    const valueWidth = contentWidth - labelWidth;
    const rowHeight = 11;

    pdf.setFontSize(12);

    rows.forEach((row) => {
      checkPage(rowHeight + 3);

      pdf.setDrawColor(0);
      pdf.setLineWidth(0.2);

      pdf.rect(marginLeft, y, labelWidth, rowHeight);
      pdf.rect(marginLeft + labelWidth, y, valueWidth, rowHeight);

      pdf.setFont("times", "bold");
      pdf.text(row.label, marginLeft + 3, y + 7);

      pdf.setFont("times", "normal");
      pdf.text(row.value || "-", marginLeft + labelWidth + 3, y + 7);

      y += rowHeight;
    });

    y += 3;
  }

  function addSignatureBlock(
    city: string,
    dateText: string,
    roleText: string,
    signerName: string
  ) {
    checkPage(55);

    const blockWidth = 65;
    const blockX = pageWidth - marginRight - blockWidth;

    y += 8;

    pdf.setFont("times", "normal");
    pdf.setFontSize(12);

    pdf.text(`${city}, ${dateText}`, blockX + blockWidth / 2, y, {
      align: "center",
    });

    y += 12;

    pdf.text(roleText, blockX + blockWidth / 2, y, {
      align: "center",
    });

    y += 6;

    if (signatureData) {
      pdf.addImage(signatureData, "PNG", blockX + 17, y, 32, 18);
    }

    y += 28;

    pdf.text(signerName, blockX + blockWidth / 2, y, {
      align: "center",
    });
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

    const tanggalDokumenFormatted = isEnglish
      ? formatTanggalInggris(tanggalDokumen)
      : formatTanggalIndonesia(tanggalDokumen);

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

    addInfoTable([
      {
        label: isEnglish ? "Time" : "Waktu",
        value: waktuKegiatan,
      },
      {
        label: isEnglish ? "Place" : "Tempat",
        value: tempatKegiatan,
      },
    ]);

    addSectionTitle(
      isEnglish ? "D. Target Participants" : "D. Sasaran Peserta"
    );
    addParagraph(sasaranPeserta);

    addSectionTitle(isEnglish ? "E. Budget Plan" : "E. Rencana Anggaran");
    addParagraph(anggaran || "-");

    addSectionTitle(isEnglish ? "F. Closing" : "F. Penutup");
    addParagraph(
      isEnglish
        ? "Thus, this activity proposal is prepared to serve as a reference for the implementation of the activity. Thank you for your attention and support."
        : "Demikian proposal kegiatan ini disusun sebagai acuan pelaksanaan kegiatan. Atas perhatian dan dukungan yang diberikan, kami ucapkan terima kasih."
    );

    addSignatureBlock(
      kota,
      tanggalDokumenFormatted,
      isEnglish ? "Person in charge," : "Penanggung Jawab,",
      namaPenanggungJawab
    );

    pdf.save(
      isEnglish
        ? safeFileName(`Activity Proposal - ${namaKegiatan}.pdf`)
        : safeFileName(`Proposal Kegiatan - ${namaKegiatan}.pdf`)
    );

    return;
  }

  if (selectedTemplateId === "invoice") {
    const namaBisnis = getValue("namaBisnis");
    const namaPelanggan = getValue("namaPelanggan");
    const nomorInvoice = getValue("nomorInvoice");
    const tanggalInvoice = getValue("tanggalInvoice");
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
      alert("Minimal isi 1 item tagihan.");
      return;
    }

    if (invoiceRows.some((item) => item.deskripsi === "")) {
      alert("Deskripsi wajib diisi pada setiap item tagihan.");
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
      alert("Jumlah dan harga satuan pada setiap item harus lebih dari 0.");
      return;
    }

    const tanggalInvoiceFormatted = isEnglish
      ? formatTanggalInggris(tanggalInvoice)
      : formatTanggalIndonesia(tanggalInvoice);

    const formatUang = (angka: number) =>
      new Intl.NumberFormat(isEnglish ? "en-US" : "id-ID", {
        style: "currency",
        currency: "IDR",
        maximumFractionDigits: 0,
      }).format(angka);

    const formatAngka = (angka: number) =>
      new Intl.NumberFormat(isEnglish ? "en-US" : "id-ID", {
        maximumFractionDigits: 2,
      }).format(angka);

    const totalTagihan = invoiceRows.reduce(
      (sum, item) => sum + item.total,
      0
    );

    const tableWidth = 132;
    const tableX = (pageWidth - tableWidth) / 2;

    function addInvoiceTextBlock(
      text: string,
      x: number,
      width: number,
      bold = false,
      size = 12,
      after = 3
    ) {
      pdf.setFont("times", bold ? "bold" : "normal");
      pdf.setFontSize(size);

      const lines = pdf.splitTextToSize(text || "-", width) as string[];

      lines.forEach((line) => {
        checkPage(8);
        pdf.text(line, x, y);
        y += 7;
      });

      y += after;
    }

    function addInvoiceRow(
      cells: {
        text: string;
        width: number;
        bold?: boolean;
        align?: "left" | "right" | "center";
      }[]
    ) {
      pdf.setFontSize(11);

      const lineHeight = 5;

      const cellLines = cells.map((cell) =>
        pdf.splitTextToSize(cell.text || " ", cell.width - 6) as string[]
      );

      const maxLines = Math.max(...cellLines.map((lines) => lines.length));
      const rowHeight = Math.max(12, maxLines * lineHeight + 7);

      checkPage(rowHeight + 3);

      let x = tableX;

      cells.forEach((cell, cellIndex) => {
        pdf.setDrawColor(0);
        pdf.setLineWidth(0.2);
        pdf.rect(x, y, cell.width, rowHeight);

        pdf.setFont("times", cell.bold ? "bold" : "normal");
        pdf.setFontSize(11);

        const lines = cellLines[cellIndex];
        const align = cell.align || "left";

        lines.forEach((line, lineIndex) => {
          const textY = y + 7 + lineIndex * lineHeight;

          if (align === "right") {
            pdf.text(line, x + cell.width - 3, textY, {
              align: "right",
            });
          } else if (align === "center") {
            pdf.text(line, x + cell.width / 2, textY, {
              align: "center",
            });
          } else {
            pdf.text(line, x + 3, textY);
          }
        });

        x += cell.width;
      });

      y += rowHeight;
    }

    y = 38;

    addCenteredText("INVOICE", 18, true, 8);

    pdf.setFont("times", "normal");
    pdf.setFontSize(12);
    pdf.text(nomorInvoice, pageWidth / 2, y, {
      align: "center",
    });

    y += 14;

    addInvoiceRow([
      {
        text: isEnglish ? "Seller / Business" : "Penjual / Bisnis",
        width: 42,
        bold: true,
      },
      {
        text: namaBisnis,
        width: 90,
      },
    ]);

    addInvoiceRow([
      {
        text: isEnglish ? "Customer" : "Pelanggan",
        width: 42,
        bold: true,
      },
      {
        text: namaPelanggan,
        width: 90,
      },
    ]);

    addInvoiceRow([
      {
        text: isEnglish ? "Invoice Date" : "Tanggal Invoice",
        width: 42,
        bold: true,
      },
      {
        text: tanggalInvoiceFormatted,
        width: 90,
      },
    ]);

    y += 12;

    addInvoiceTextBlock(
      isEnglish ? "Billing Details" : "Rincian Tagihan",
      tableX,
      tableWidth,
      true,
      13,
      4
    );

    addInvoiceRow([
      {
        text: isEnglish ? "Description" : "Deskripsi",
        width: 53,
        bold: true,
      },
      {
        text: isEnglish ? "Qty" : "Jumlah",
        width: 20,
        bold: true,
      },
      {
        text: isEnglish ? "Unit Price" : "Harga Satuan",
        width: 30,
        bold: true,
      },
      {
        text: "Total",
        width: 29,
        bold: true,
      },
    ]);

    invoiceRows.forEach((item) => {
      addInvoiceRow([
        {
          text: item.deskripsi,
          width: 53,
        },
        {
          text: formatAngka(item.jumlah),
          width: 20,
        },
        {
          text: formatUang(item.hargaSatuan),
          width: 30,
        },
        {
          text: formatUang(item.total),
          width: 29,
        },
      ]);
    });

    addInvoiceRow([
      {
        text: " ",
        width: 53,
      },
      {
        text: " ",
        width: 20,
      },
      {
        text: isEnglish ? "Grand Total" : "Total Tagihan",
        width: 30,
        bold: true,
      },
      {
        text: formatUang(totalTagihan),
        width: 29,
        bold: true,
      },
    ]);

    if (catatan) {
      y += 12;

      addInvoiceTextBlock(
        isEnglish ? "Payment Notes" : "Catatan Pembayaran",
        tableX,
        tableWidth,
        true,
        13,
        4
      );

      addInvoiceTextBlock(catatan, tableX, tableWidth, false, 12, 4);
    }

        y += 6;

    addInvoiceTextBlock(
      isEnglish
        ? "Thank you for your trust and cooperation."
        : "Terima kasih atas kepercayaan dan kerja samanya.",
      tableX,
      tableWidth,
      false,
      12,
      0
    );

    const signatureBlockWidth = 60;
    const signatureBlockX = pageWidth - marginRight - signatureBlockWidth;

    let signatureY = Math.max(y + 12, 205);

    if (signatureY + 50 > pageHeight - 18) {
      if (y < 220) {
        signatureY = pageHeight - 68;
      } else {
        pdf.addPage();
        signatureY = 45;
      }
    }

    pdf.setFont("times", "normal");
    pdf.setFontSize(12);

    pdf.text(
      isEnglish ? "Seller," : "Penjual,",
      signatureBlockX + signatureBlockWidth / 2,
      signatureY,
      {
        align: "center",
      }
    );

    if (signatureData) {
      pdf.addImage(
        signatureData,
        "PNG",
        signatureBlockX + 14,
        signatureY + 8,
        34,
        20
      );
    }

    pdf.text(
      namaBisnis,
      signatureBlockX + signatureBlockWidth / 2,
      signatureY + 42,
      {
        align: "center",
      }
    );

    pdf.save(safeFileName(`Invoice - ${nomorInvoice}.pdf`));
    return;
  }

  // Fallback sementara untuk template lain
  const title = selectedTemplate?.name || "Dokumen";

  addCenteredText(title.toUpperCase(), 16, true, 12);

  selectedTemplate?.fields.forEach((field) => {
    let value = getValue(field.name);

    if (field.type === "date") {
      value = isEnglish
        ? formatTanggalInggris(value)
        : formatTanggalIndonesia(value);
    }

    addParagraph(`${field.label}: ${value || "-"}`);
  });

  addSignatureBlock(
    getValue("kota") || "-",
    "",
    isEnglish ? "Signature," : "Tanda tangan,",
    getValue("nama") ||
      getValue("namaPenandatangan") ||
      getValue("namaPenanggungJawab") ||
      getValue("namaBisnis") ||
      getValue("notulis") ||
      getValue("namaPemberi") ||
      "Nama Penandatangan"
  );

  pdf.save(
    safeFileName(`${title} - ${new Date().getTime()}.pdf`)
  );
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
  alert("Tanda tangan wajib diisi.");
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
  alert("Download Word belum tersedia untuk template ini.");
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
    alert("Tanggal selesai izin tidak boleh lebih awal dari tanggal mulai izin.");
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
    alert("Minimal isi 1 entry pembahasan rapat.");
    return;
  }

  if (pembahasanRows.some((entry) => entry.pembahasan === "")) {
    alert("Kolom Pembahasan wajib diisi pada setiap entry yang digunakan.");
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
    alert("Minimal isi 1 item tagihan.");
    return;
  }

  if (invoiceRows.some((item) => item.deskripsi === "")) {
    alert("Deskripsi wajib diisi pada setiap item tagihan.");
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
    alert("Jumlah dan harga satuan pada setiap item harus lebih dari 0.");
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
    new Intl.NumberFormat(isEnglish ? "en-US" : "id-ID", {
      style: "currency",
      currency: "IDR",
      maximumFractionDigits: 0,
    }).format(angka);

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
          <p className="mb-3 text-sm font-semibold uppercase tracking-wide text-blue-600">
            DokumenJadi
          </p>

          <h1 className="mb-4 max-w-3xl text-4xl font-bold tracking-tight md:text-5xl">
            Buat dokumen otomatis dalam hitungan detik.
          </h1>

          <p className="max-w-2xl text-lg text-slate-600">
            Pilih template, isi form sederhana, lalu download dokumen yang rapi
            dan siap digunakan.
          </p>
        </div>

        <div className="grid gap-6 lg:grid-cols-[1fr_1.2fr]">
          <div>
            <h2 className="mb-4 text-2xl font-bold">Pilih Template</h2>

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
                      {template.name}
                    </h3>
                    <p className="text-sm leading-6 text-slate-600">
                      {template.description}
                    </p>
                  </button>
                );
              })}
            </div>
          </div>

          <div className="rounded-3xl bg-white p-6 shadow-sm">
            <h2 className="mb-2 text-2xl font-bold">
              {selectedTemplate?.name}
            </h2>

            <p className="mb-6 text-slate-600">
              Isi data berikut untuk membuat dokumen.
            </p>

            <form ref={formRef} onSubmit={handleSubmit} className="grid gap-4">
              {selectedTemplate?.fields.map((field) => (
                <div key={field.name}>
                  <label className="mb-2 block text-sm font-medium text-slate-700">
                    {field.label}
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
          Daftar Pembahasan Rapat
        </h3>
        <p className="text-sm text-slate-500">
          Tambahkan pembahasan, keputusan, dan tindak lanjut sesuai kebutuhan.
        </p>
      </div>

      <button
        type="button"
        onClick={addNotulenEntry}
        className="rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-700"
      >
        + Tambah Entry
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
              Entry {index + 1}
            </h4>

            {notulenEntries.length > 1 && (
              <button
                type="button"
                onClick={() => removeNotulenEntry(index)}
                className="rounded-lg border border-red-200 px-3 py-1 text-sm font-semibold text-red-600 hover:bg-red-50"
              >
                Hapus
              </button>
            )}
          </div>

          <div className="grid gap-3">
            <div>
              <label className="mb-2 block text-sm font-medium text-slate-700">
                Pembahasan <span className="text-red-500">*</span>
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
                Keputusan
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
                Tindak Lanjut
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
  <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
    <div className="mb-4 flex items-center justify-between gap-3">
      <div>
        <h3 className="text-lg font-semibold text-slate-900">
          Daftar Item Tagihan
        </h3>
        <p className="text-sm text-slate-500">
          Tambahkan item, jumlah, dan harga satuan sesuai kebutuhan invoice.
        </p>
      </div>

      <button
        type="button"
        onClick={addInvoiceItem}
        className="rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-700"
      >
        + Tambah Item
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
              Item {index + 1}
            </h4>

            {invoiceItems.length > 1 && (
              <button
                type="button"
                onClick={() => removeInvoiceItem(index)}
                className="rounded-lg border border-red-200 px-3 py-1 text-sm font-semibold text-red-600 hover:bg-red-50"
              >
                Hapus
              </button>
            )}
          </div>

          <div className="grid gap-3">
            <div>
              <label className="mb-2 block text-sm font-medium text-slate-700">
                Deskripsi Produk/Jasa <span className="text-red-500">*</span>
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
                  Jumlah <span className="text-red-500">*</span>
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
                  Harga Satuan <span className="text-red-500">*</span>
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
                  Bahasa Dokumen <span className="text-red-500">*</span>
                </label>

                <select
                  name="bahasaDokumen"
                  defaultValue="id"
                  required
                  className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none focus:border-blue-500"
                >
                  <option value="id">Indonesia</option>
                  <option value="en">Inggris</option>
                </select>
              </div>

<div>
  <label className="mb-2 block text-sm font-medium text-slate-700">
    Tanda Tangan <span className="text-red-500">*</span>
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
  Buat tanda tangan langsung pada area di atas, lalu pastikan posisinya berada di tengah area.
</p>

    <button
      type="button"
      onClick={clearSignature}
      className="rounded-lg border border-slate-300 px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50"
    >
      Hapus Tanda Tangan
    </button>
  </div>
</div>

              <div className="mt-4 grid gap-3 sm:grid-cols-2">
  <button
    type="button"
    onClick={() => setShowClearConfirm(true)}
    className="rounded-xl border border-slate-300 px-5 py-3 font-semibold text-slate-700 transition hover:bg-slate-50"
  >
    Kosongkan Isian
  </button>

  <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
    <button
      type="submit"
      name="downloadType"
      value="word"
      className="rounded-xl bg-blue-600 px-5 py-3 font-semibold text-white transition hover:bg-blue-700"
    >
      Download Word
    </button>

    <button
      type="submit"
      name="downloadType"
      value="pdf"
      className="rounded-xl bg-red-600 px-5 py-3 font-semibold text-white transition hover:bg-red-700"
    >
      Download PDF
    </button>
  </div>
</div>
</form>
            {showClearConfirm && (
  <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4">
    <div className="w-full max-w-md rounded-2xl bg-white p-6 shadow-xl">
      <h3 className="mb-2 text-xl font-bold text-slate-900">
        Kosongkan Semua Isian?
      </h3>

      <p className="mb-6 text-sm leading-6 text-slate-600">
        Semua data yang sudah diisi, termasuk tanda tangan, akan dikosongkan.
        Apakah Anda yakin ingin melanjutkan?
      </p>

      <div className="grid gap-3 sm:grid-cols-2">
        <button
          type="button"
          onClick={() => setShowClearConfirm(false)}
          className="rounded-xl border border-slate-300 px-4 py-3 font-semibold text-slate-700 hover:bg-slate-50"
        >
          Tidak
        </button>

        <button
          type="button"
          onClick={clearAllInputs}
          className="rounded-xl bg-red-600 px-4 py-3 font-semibold text-white hover:bg-red-700"
        >
          Ya, Kosongkan
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