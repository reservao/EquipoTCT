import { NextRequest, NextResponse } from "next/server";
import { parseCompetenciaDocx } from "@/lib/docx-parser";
import { generateSotWorkbook, buildOutputFileName } from "@/lib/sot-generator";

export const runtime = "nodejs";

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof File)) {
      return NextResponse.json({ error: "No se recibió ningún archivo." }, { status: 400 });
    }
    if (!file.name.toLowerCase().endsWith(".docx")) {
      return NextResponse.json(
        { error: "El archivo debe ser un Word (.docx)." },
        { status: 400 }
      );
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const data = await parseCompetenciaDocx(buffer);
    const workbookBuffer = await generateSotWorkbook(data);
    const fileName = buildOutputFileName(data);

    return NextResponse.json({
      fileName,
      fileBase64: workbookBuffer.toString("base64"),
      summary: {
        nombre: data.nombre,
        perfiles: data.perfiles,
        gerencia: data.gerencia,
        superintendencia: data.superintendencia,
        area: data.area,
        codigo: data.codigo,
        activities: data.activities.map((a) => ({
          title: a.title,
          criteriosCount: a.criterios.length,
        })),
      },
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Error desconocido al procesar el archivo.";
    return NextResponse.json({ error: message }, { status: 422 });
  }
}
