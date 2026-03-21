import React, { useMemo, useState } from "react";
import { motion } from "framer-motion";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Badge } from "@/components/ui/badge";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { Tabs, TabsList, TabsTrigger, TabsContent } from "@/components/ui/tabs";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Plus, Download, Trash2, Search, LayoutGrid, Table2, CalendarRange, Briefcase, Landmark } from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, CartesianGrid, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";

const MONTHS = Array.from({ length: 12 }, (_, i) => i + 1);

const INITIAL_DATA = [
  {
    id: "gobernanza",
    name: "Gobernanza",
    items: [
      {
        id: "g-1",
        subcategory: "Convenio / Tratado Marco",
        title:
          "Contratar consultor abogado (Alan Thomson) para acompañamiento del BID en la estructuración técnica de la gobernanza",
        months: [2],
      },
      {
        id: "g-2",
        subcategory: "Convenio / Tratado Marco",
        title:
          "Análisis sobre si corresponde convenio o tratado marco y modelo de gobernanza para proyectos piloto",
        months: [3],
      },
      {
        id: "g-3",
        subcategory: "Convenio / Tratado Marco",
        title:
          "Preparación y socialización del borrador de convenio/tratado con comisiones y/o gobernanza para pilotos",
        months: [4, 5],
      },
      {
        id: "g-4",
        subcategory: "Convenio / Tratado Marco",
        title:
          "Firma del borrador de convenio/tratado marco en consejo intersectorial de ministros",
        months: [6],
      },
      {
        id: "g-5",
        subcategory: "Convenio / Tratado Marco",
        title:
          "Instrumento legal borrador para la gobernanza de los pilotos (a nivel de ministros de los 2 países participantes)",
        months: [6],
      },
      {
        id: "g-6",
        subcategory: "Unidad Ejecutora Inicial",
        title:
          "Contratar 2 consultores (coordinador y apoyo técnico) dentro de SIECA temporalmente",
        months: [7, 8, 9, 10, 11, 12],
      },
      {
        id: "g-7",
        subcategory: "Entrada en funcionamiento de instituciones regionales nuevas",
        title:
          "Diseño detallado del esquema administrativo, operativo, financiero y funcional conforme al CM y estudio técnico",
        months: [7, 8, 9, 10, 11, 12],
      },
      {
        id: "g-8",
        subcategory: "Entrada en funcionamiento de instituciones regionales nuevas",
        title: "Inicio en funcionamiento",
        months: [],
      },
    ],
  },
  {
    id: "tecnicos",
    name: "Técnicos",
    items: [
      {
        id: "t-1",
        subcategory: "Estudios de ingeniería",
        title:
          "Preparación de TdR para el diseño de los proyectos piloto (PN-CR Paso Canoas y ES-HO Amatillo) y estudio de factibilidad técnica avanzada de Cargo Pass",
        months: [1, 2],
      },
      {
        id: "t-2",
        subcategory: "Estudios de ingeniería",
        title:
          "Proyectos piloto (PN-CR Paso Canoas y ES-HO Amatillo) y estudio de factibilidad técnica avanzada de Cargo Pass",
        months: [3, 4, 5, 6, 7, 8, 9, 10],
      },
      {
        id: "t-3",
        subcategory: "Estudios de derivación de carga y viabilidad",
        title:
          "Estudio de demanda, DAP y viabilidad de derivación de carga al CargoPass",
        months: [2, 3, 4, 5, 6, 7, 8, 9],
      },
      {
        id: "t-4",
        subcategory: "Análisis económica, financiera, riesgos y legal",
        title:
          "Estudio de viabilidad financiera y económica final, análisis de riesgos, acorde a procesos y guías de programación en los países",
        months: [9, 10, 11],
      },
      {
        id: "t-5",
        subcategory: "Estudios de tiempos de viaje",
        title:
          "Estudios de tiempos de viaje de todo el corredor, incluyendo puestos de frontera",
        months: [3, 4, 5, 6, 7, 8],
      },
      {
        id: "t-6",
        subcategory: "Análisis del OEA para todo el corredor",
        title:
          "Diagnóstico para el análisis de los programas nacionales del Operador Económico Autorizado (tipificación, alcance y acuerdos de reconocimiento mutuo)",
        months: [2, 3, 4, 5, 6, 7, 8, 9, 10],
      },
    ],
  },
  {
    id: "fondeo",
    name: "Fondeo",
    items: [
      {
        id: "f-1",
        subcategory: "Fondo 1 de preinversión",
        title:
          "Para Unidad Ejecutora inicial y para estudios de preinversión (CTs actuales del Banco, nuevos fondos JICA, UE, etc.)",
        months: [3, 4, 5, 6, 7, 8, 9, 10, 11],
      },
      {
        id: "f-2",
        subcategory: "Fondo 2 de operación para la gobernanza",
        title:
          "Para Unidad Ejecutora inicial y arranque del funcionamiento de la gobernanza (CTs actuales del Banco, nuevos fondos JICA, UE, etc.)",
        months: [],
      },
      {
        id: "f-3",
        subcategory: "Fondo 3 implementación",
        title: "CCLIP del BID más un champion",
        months: [],
      },
    ],
  },
];

const SECTION_STYLES = {
  gobernanza: { icon: Landmark, color: "#003F72", header: "bg-[#003F72] text-white", chip: "bg-[#EAF3F9] text-[#003F72] border-[#C6DCEB]" },
  tecnicos: { icon: Briefcase, color: "#0F5B78", header: "bg-[#0F5B78] text-white", chip: "bg-[#ECFEFF] text-[#155E75] border-[#A5F3FC]" },
  fondeo: { icon: CalendarRange, color: "#5B6B7A", header: "bg-[#5B6B7A] text-white", chip: "bg-[#F1F5F9] text-[#334155] border-[#CBD5E1]" },
};

function uid() {
  return Math.random().toString(36).slice(2, 9);
}

function exportExcel(data) {
  const rows = [];
  data.forEach((section) => {
    section.items.forEach((item) => {
      const row = {
        Sección: section.name,
        Subcategoría: item.subcategory,
        Actividad: item.title,
      };
      MONTHS.forEach((month) => {
        row[`Mes ${month}`] = item.months.includes(month) ? 1 : "";
      });
      rows.push(row);
    });
  });

  const worksheet = XLSX.utils.json_to_sheet(rows);
  worksheet["!cols"] = [{ wch: 22 }, { wch: 34 }, { wch: 90 }, ...Array.from({ length: 12 }, () => ({ wch: 10 }))];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Actividades_Hitos");
  XLSX.writeFile(workbook, "actividades_hitos.xlsx");
}

function getRange(months) {
  if (!months.length) return null;
  return { start: Math.min(...months), end: Math.max(...months) };
}

function getDuration(months) {
  const range = getRange(months);
  return range ? range.end - range.start + 1 : 0;
}

function GanttBar({ months, onToggle, color, critical }) {
  const range = getRange(months);
  return (
    <div className="relative h-10 w-full min-w-[420px]">
      <div className="absolute inset-0 grid grid-cols-12">
        {MONTHS.map((m) => (
          <div key={m} className="border-r border-[#E6EDF3] last:border-r-0" />
        ))}
      </div>

      <div className="absolute left-0 right-0 top-1/2 h-px -translate-y-1/2 bg-[#C8D4DF]" />

      {range && (
        <div
          className={`absolute top-1/2 h-5 -translate-y-1/2 rounded-full shadow-sm ${critical ? "ring-2 ring-[#FECACA]" : ""}`}
          style={{
            left: `calc(${((range.start - 1) / 12) * 100}% + 6px)`,
            width: `calc(${((range.end - range.start + 1) / 12) * 100}% - 12px)`,
            backgroundColor: critical ? "#C81E1E" : color,
          }}
        />
      )}

      <div className="absolute inset-0 grid grid-cols-12">
        {MONTHS.map((month) => (
          <button key={month} onClick={() => onToggle(month)} className="h-full w-full" title={`Mes ${month}`} />
        ))}
      </div>
    </div>
  );
}

export default function MiniAppActividadesHitos() {
  const [sections, setSections] = useState(INITIAL_DATA);
  const [query, setQuery] = useState("");
  const [view, setView] = useState("cronograma");
  const [draft, setDraft] = useState({
    sectionId: "gobernanza",
    subcategory: "",
    title: "",
    months: [],
  });

  const criticalIds = useMemo(() => {
    const all = sections.flatMap((section) => section.items);
    const maxDuration = Math.max(0, ...all.map((item) => getDuration(item.months)));
    return new Set(all.filter((item) => getDuration(item.months) === maxDuration && maxDuration > 0).map((item) => item.id));
  }, [sections]);

  const stats = useMemo(() => {
    const all = sections.flatMap((s) => s.items);
    return {
      total: all.length,
      active: all.filter((i) => i.months.length > 0).length,
      totalMarks: all.reduce((acc, i) => acc + i.months.length, 0),
    };
  }, [sections]);

  const monthlyCounts = useMemo(() => {
    const counts = Array.from({ length: 12 }, (_, i) => ({ mes: i + 1, total: 0 }));
    sections.forEach((section) => {
      section.items.forEach((item) => {
        item.months.forEach((m) => {
          counts[m - 1].total += 1;
        });
      });
    });
    return counts;
  }, [sections]);

  const filteredSections = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return sections;
    return sections
      .map((section) => ({
        ...section,
        items: section.items.filter(
          (item) =>
            section.name.toLowerCase().includes(q) ||
            item.subcategory.toLowerCase().includes(q) ||
            item.title.toLowerCase().includes(q)
        ),
      }))
      .filter((section) => section.items.length > 0);
  }, [sections, query]);

  const toggleMonth = (sectionId, itemId, month) => {
    setSections((prev) =>
      prev.map((section) =>
        section.id !== sectionId
          ? section
          : {
              ...section,
              items: section.items.map((item) =>
                item.id !== itemId
                  ? item
                  : {
                      ...item,
                      months: item.months.includes(month)
                        ? item.months.filter((m) => m !== month)
                        : [...item.months, month].sort((a, b) => a - b),
                    }
              ),
            }
      )
    );
  };

  const updateItem = (sectionId, itemId, patch) => {
    setSections((prev) =>
      prev.map((section) =>
        section.id !== sectionId
          ? section
          : {
              ...section,
              items: section.items.map((item) => (item.id === itemId ? { ...item, ...patch } : item)),
            }
      )
    );
  };

  const deleteItem = (sectionId, itemId) => {
    setSections((prev) =>
      prev.map((section) =>
        section.id !== sectionId
          ? section
          : { ...section, items: section.items.filter((item) => item.id !== itemId) }
      )
    );
  };

  const addItem = () => {
    if (!draft.title.trim() || !draft.subcategory.trim()) return;
    setSections((prev) =>
      prev.map((section) =>
        section.id !== draft.sectionId
          ? section
          : {
              ...section,
              items: [
                ...section.items,
                {
                  id: uid(),
                  subcategory: draft.subcategory,
                  title: draft.title,
                  months: [...draft.months].sort((a, b) => a - b),
                },
              ],
            }
      )
    );
    setDraft({ sectionId: draft.sectionId, subcategory: "", title: "", months: [] });
  };

  return (
    <div className="min-h-screen bg-[#F4F7F9] text-[#1F2A37]">
      <div className="border-b border-[#D8E1E8] bg-white">
        <div className="mx-auto max-w-7xl px-4 py-5 md:px-8">
          <div className="flex items-center justify-between gap-4">
            <div className="flex items-center gap-4">
              <div className="h-10 w-1.5 rounded-full bg-[#003F72]" />
              <div>
                <p className="text-xs font-semibold uppercase tracking-[0.18em] text-[#5B6B7A]">Banco Interamericano de Desarrollo</p>
                <h1 className="text-2xl font-semibold tracking-tight text-[#003F72] md:text-3xl">Actividades / Hitos</h1>
              </div>
            </div>
            <div className="hidden text-right md:block">
              <p className="text-sm text-[#5B6B7A]">Cronograma editable</p>
              <p className="text-xs text-[#7B8794]">Vista Gantt con ruta crítica</p>
            </div>
          </div>
        </div>
      </div>

      <div className="mx-auto max-w-7xl space-y-6 px-4 py-6 md:px-8">
        <motion.div initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }}>
          <Card className="rounded-xl border border-[#D8E1E8] bg-white shadow-sm">
            <CardHeader className="pb-4">
              <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
                <div>
                  <CardTitle className="text-xl font-semibold text-[#003F72]">Panel de control</CardTitle>
                  <p className="mt-1 text-sm text-[#5B6B7A]">Edita actividades, activa meses, agrega filas y exporta el cronograma en Excel.</p>
                </div>
                <div className="flex flex-wrap gap-2">
                  <Button variant="outline" className="rounded-md border-[#C8D4DF] text-[#003F72]" onClick={() => exportExcel(sections)}>
                    <Download className="mr-2 h-4 w-4" /> Exportar en Excel
                  </Button>
                </div>
              </div>
            </CardHeader>
            <CardContent>
              <div className="grid gap-3 md:grid-cols-4">
                <Card className="rounded-lg border border-[#D8E1E8] shadow-none"><CardContent className="p-4"><p className="text-xs font-semibold uppercase tracking-wide text-[#7B8794]">Total de actividades</p><p className="mt-2 text-2xl font-semibold text-[#003F72]">{stats.total}</p></CardContent></Card>
                <Card className="rounded-lg border border-[#D8E1E8] shadow-none"><CardContent className="p-4"><p className="text-xs font-semibold uppercase tracking-wide text-[#7B8794]">Con meses asignados</p><p className="mt-2 text-2xl font-semibold text-[#003F72]">{stats.active}</p></CardContent></Card>
                <Card className="rounded-lg border border-[#D8E1E8] shadow-none"><CardContent className="p-4"><p className="text-xs font-semibold uppercase tracking-wide text-[#7B8794]">Hitos marcados</p><p className="mt-2 text-2xl font-semibold text-[#003F72]">{stats.totalMarks}</p></CardContent></Card>
                <Card className="rounded-lg border border-[#D8E1E8] shadow-none"><CardContent className="p-4"><div className="flex items-center gap-2"><Search className="h-4 w-4 text-[#7B8794]" /><Input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar actividad o categoría" className="border-0 p-0 text-[#1F2A37] shadow-none focus-visible:ring-0" /></div></CardContent></Card>
              </div>
            </CardContent>
          </Card>
        </motion.div>

        <Card className="rounded-xl border border-[#D8E1E8] bg-white shadow-sm">
          <CardHeader>
            <CardTitle className="text-[#003F72]">Carga de actividades por mes</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="h-64 w-full">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={monthlyCounts}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="mes" />
                  <YAxis />
                  <Tooltip />
                  <Bar dataKey="total" />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </CardContent>
        </Card>

        <Tabs value={view} onValueChange={setView} className="w-full">
          <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <TabsList className="grid w-full max-w-sm grid-cols-2 rounded-md bg-[#EAF0F5]">
              <TabsTrigger value="cronograma" className="rounded-md data-[state=active]:bg-white data-[state=active]:text-[#003F72]"><LayoutGrid className="mr-2 h-4 w-4" /> Cronograma</TabsTrigger>
              <TabsTrigger value="lista" className="rounded-md data-[state=active]:bg-white data-[state=active]:text-[#003F72]"><Table2 className="mr-2 h-4 w-4" /> Lista</TabsTrigger>
            </TabsList>

            <Dialog>
              <DialogTrigger asChild>
                <Button className="rounded-md bg-[#003F72] hover:bg-[#00345E]"><Plus className="mr-2 h-4 w-4" /> Agregar actividad</Button>
              </DialogTrigger>
              <DialogContent className="rounded-xl border border-[#D8E1E8] sm:max-w-2xl">
                <DialogHeader>
                  <DialogTitle className="text-[#003F72]">Nueva actividad</DialogTitle>
                </DialogHeader>
                <div className="grid gap-4">
                  <div className="grid gap-2 md:grid-cols-2">
                    <div>
                      <p className="mb-2 text-sm font-medium text-[#334155]">Sección</p>
                      <Select value={draft.sectionId} onValueChange={(v) => setDraft((d) => ({ ...d, sectionId: v }))}>
                        <SelectTrigger className="rounded-md border-[#CBD5E1]"><SelectValue /></SelectTrigger>
                        <SelectContent>{sections.map((s) => <SelectItem key={s.id} value={s.id}>{s.name}</SelectItem>)}</SelectContent>
                      </Select>
                    </div>
                    <div>
                      <p className="mb-2 text-sm font-medium text-[#334155]">Subcategoría</p>
                      <Input value={draft.subcategory} onChange={(e) => setDraft((d) => ({ ...d, subcategory: e.target.value }))} placeholder="Ej. Convenio / Tratado Marco" className="rounded-md border-[#CBD5E1]" />
                    </div>
                  </div>
                  <div>
                    <p className="mb-2 text-sm font-medium text-[#334155]">Actividad</p>
                    <Textarea value={draft.title} onChange={(e) => setDraft((d) => ({ ...d, title: e.target.value }))} placeholder="Describe la actividad o hito" className="min-h-[120px] rounded-md border-[#CBD5E1]" />
                  </div>
                  <div>
                    <p className="mb-2 text-sm font-medium text-[#334155]">Meses</p>
                    <div className="grid grid-cols-6 gap-2 md:grid-cols-12">
                      {MONTHS.map((month) => {
                        const active = draft.months.includes(month);
                        return (
                          <Button key={month} type="button" variant={active ? "default" : "outline"} className={`rounded-md ${active ? "bg-[#003F72] hover:bg-[#00345E]" : "border-[#CBD5E1] text-[#334155]"}`} onClick={() => setDraft((d) => ({ ...d, months: active ? d.months.filter((m) => m !== month) : [...d.months, month] }))}>{month}</Button>
                        );
                      })}
                    </div>
                  </div>
                  <div className="flex justify-end"><Button onClick={addItem} className="rounded-md bg-[#003F72] hover:bg-[#00345E]">Guardar actividad</Button></div>
                </div>
              </DialogContent>
            </Dialog>
          </div>

          <TabsContent value="cronograma" className="mt-6">
            <Card className="rounded-xl border border-[#D8E1E8] bg-white shadow-sm">
              <CardContent className="p-4">
                <div className="mb-4 flex flex-wrap items-center gap-4 text-sm text-[#52606D]">
                  <div className="flex items-center gap-2"><span className="inline-block h-3 w-8 rounded-full bg-[#003F72]" /> Actividades</div>
                  <div className="flex items-center gap-2"><span className="inline-block h-3 w-8 rounded-full bg-[#C81E1E]" /> Ruta crítica</div>
                </div>
                <ScrollArea className="w-full whitespace-nowrap">
                  <div className="min-w-[1350px]">
                    <div className="mb-4 grid grid-cols-[220px_420px_120px_1fr] rounded-lg border border-[#D8E1E8] bg-[#F8FAFC] text-sm font-semibold text-[#52606D]">
                      <div className="border-r border-[#D8E1E8] px-4 py-3">Componente</div>
                      <div className="border-r border-[#D8E1E8] px-4 py-3">Actividad</div>
                      <div className="border-r border-[#D8E1E8] px-4 py-3 text-center">Periodo</div>
                      <div className="grid grid-cols-12">
                        {MONTHS.map((month) => <div key={month} className="border-r border-[#D8E1E8] px-2 py-3 text-center last:border-r-0">{month}</div>)}
                      </div>
                    </div>

                    {filteredSections.map((section) => {
                      const style = SECTION_STYLES[section.id];
                      const Icon = style.icon;
                      return (
                        <div key={section.id} className="mb-6">
                          <div className={`mb-2 flex items-center gap-3 rounded-lg px-4 py-3 ${style.header}`}>
                            <Icon className="h-5 w-5" />
                            <span className="text-base font-semibold">{section.name}</span>
                          </div>

                          <div className="overflow-hidden rounded-lg border border-[#D8E1E8]">
                            {section.items.map((item) => {
                              const range = getRange(item.months);
                              const isCritical = criticalIds.has(item.id);
                              return (
                                <div key={item.id} className="grid grid-cols-[220px_420px_120px_1fr] border-b border-[#E7EDF2] bg-white last:border-b-0">
                                  <div className="border-r border-[#E7EDF2] px-4 py-4">
                                    <Badge className={`whitespace-normal border px-2 py-1 text-left font-medium ${style.chip}`}>{item.subcategory}</Badge>
                                  </div>
                                  <div className="border-r border-[#E7EDF2] px-4 py-3">
                                    <Textarea value={item.title} onChange={(e) => updateItem(section.id, item.id, { title: e.target.value })} className="min-h-[82px] rounded-md border-[#D6DEE5] bg-[#FBFDFE] text-sm shadow-none focus-visible:ring-1 focus-visible:ring-[#9FB6C8]" />
                                  </div>
                                  <div className="flex items-center justify-center border-r border-[#E7EDF2] px-3 py-3 text-sm font-medium text-[#486170]">{range ? `${range.start}–${range.end}` : "—"}</div>
                                  <div className="px-3 py-3">
                                    <GanttBar months={item.months} color={style.color} critical={isCritical} onToggle={(month) => toggleMonth(section.id, item.id, month)} />
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </ScrollArea>
              </CardContent>
            </Card>
          </TabsContent>

          <TabsContent value="lista" className="mt-6">
            <div className="grid gap-4">
              {filteredSections.map((section) => {
                const style = SECTION_STYLES[section.id];
                const Icon = style.icon;
                return (
                  <Card key={section.id} className="rounded-xl border border-[#D8E1E8] bg-white shadow-sm">
                    <CardHeader className="border-b border-[#EEF2F6] pb-4">
                      <CardTitle className="flex items-center gap-2 text-[#003F72]"><Icon className="h-5 w-5" /> {section.name}</CardTitle>
                    </CardHeader>
                    <CardContent className="space-y-3 pt-5">
                      {section.items.map((item) => (
                        <motion.div key={item.id} initial={{ opacity: 0, y: 4 }} animate={{ opacity: 1, y: 0 }} className="rounded-lg border border-[#D8E1E8] bg-white p-4">
                          <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                            <div className="space-y-2 flex-1">
                              <Badge className={`border px-2 py-1 font-medium ${style.chip}`}>{item.subcategory}</Badge>
                              <Textarea value={item.title} onChange={(e) => updateItem(section.id, item.id, { title: e.target.value })} className="min-h-[88px] rounded-md border-[#D6DEE5]" />
                              <div className="flex flex-wrap gap-2">
                                {MONTHS.map((month) => (
                                  <Button key={month} type="button" size="sm" variant={item.months.includes(month) ? "default" : "outline"} className={`rounded-md ${item.months.includes(month) ? "bg-[#003F72] hover:bg-[#00345E]" : "border-[#CBD5E1] text-[#334155]"}`} onClick={() => toggleMonth(section.id, item.id, month)}>{month}</Button>
                                ))}
                              </div>
                            </div>
                            <Button variant="outline" size="icon" className="rounded-md border-[#CBD5E1] text-[#7A1F1F]" onClick={() => deleteItem(section.id, item.id)}><Trash2 className="h-4 w-4" /></Button>
                          </div>
                        </motion.div>
                      ))}
                    </CardContent>
                  </Card>
                );
              })}
            </div>
          </TabsContent>
        </Tabs>
      </div>
    </div>
  );
}
