import Link from "next/link";

const modules = [
  {
    href: "/sot",
    title: "SOT — Set de Observación en Terreno",
    description:
      "Sube el Word de una Competencia (UCL) y genera automáticamente el Excel del SOT, listo para usar en terreno.",
    status: "Disponible" as const,
  },
  {
    href: null,
    title: "Próximo módulo",
    description: "Cuéntame qué otra tarea del equipo TCT quieres automatizar y la agregamos aquí.",
    status: "Próximamente" as const,
  },
];

export default function Home() {
  return (
    <div className="mx-auto max-w-5xl px-6 py-10">
      <h1 className="text-2xl font-semibold tracking-tight">Herramientas del equipo TCT</h1>
      <p className="mt-2 text-sm text-slate-500">
        Automatizaciones para Talentos y Capacidades del Trabajo, construidas paso a paso.
      </p>

      <div className="mt-8 grid gap-4 sm:grid-cols-2">
        {modules.map((m) => {
          const content = (
            <div
              className={`h-full rounded-xl border border-slate-200 bg-white p-5 shadow-sm transition ${
                m.href ? "hover:border-slate-300 hover:shadow-md" : "opacity-60"
              }`}
            >
              <div className="flex items-center justify-between">
                <h2 className="font-medium">{m.title}</h2>
                <span
                  className={`rounded-full px-2 py-0.5 text-xs font-medium ${
                    m.status === "Disponible"
                      ? "bg-emerald-100 text-emerald-700"
                      : "bg-slate-100 text-slate-500"
                  }`}
                >
                  {m.status}
                </span>
              </div>
              <p className="mt-2 text-sm text-slate-500">{m.description}</p>
            </div>
          );
          return m.href ? (
            <Link key={m.title} href={m.href}>
              {content}
            </Link>
          ) : (
            <div key={m.title}>{content}</div>
          );
        })}
      </div>
    </div>
  );
}
