# Resultados-Simulacro-Saber-Simón Bolívar 
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Saber 11° - I.E. Simón Bolívar</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Chart.js CDN -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- SheetJS (XLSX) CDN -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; background-color: #f8fafc; }
        .glass-card {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(226, 232, 240, 0.8);
        }
    </style>
</head>
<body class="text-slate-800 min-h-screen flex flex-col">

    <!-- HEADER / NAVIGATION -->
    <header class="bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 text-white shadow-xl">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
            <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
                <div class="flex items-center space-x-4">
                    <div class="bg-white/10 p-3 rounded-2xl backdrop-blur-md border border-white/20">
                        <i class="fa-solid fa-graduation-cap text-3xl text-blue-400"></i>
                    </div>
                    <div>
                        <h1 class="text-2xl sm:text-3xl font-extrabold tracking-tight">Institución Educativa Simón Bolívar</h1>
                        <p class="text-blue-200 text-sm font-medium mt-0.5">Resultados Saber 11° — Oficial Institución Educativa Simón Bolívar</p>
                    </div>
                </div>
                
                <div class="flex flex-wrap items-center gap-3">
                    <div class="bg-white/10 backdrop-blur-md rounded-xl p-1.5 border border-white/20 flex items-center space-x-2">
                        <label for="gradeFilter" class="text-xs font-semibold text-blue-200 uppercase pl-2"><i class="fa-solid fa-filter mr-1"></i> Grado:</label>
                        <select id="gradeFilter" onchange="filterData()" class="bg-slate-800 text-white text-sm rounded-lg px-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-blue-400 border border-slate-700">
                            <option value="ALL">Todos los Grados</option>
                            <option value="11°" selected>Grado 11°</option>
                        </select>
                    </div>

                    <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-500 text-white font-semibold text-sm px-4 py-2 rounded-xl shadow-lg hover:shadow-emerald-900/30 transition-all duration-200 flex items-center gap-2 border border-emerald-400/30">
                        <i class="fa-solid fa-file-excel text-base"></i>
                        <span>Exportar Excel</span>
                    </button>
                </div>
            </div>
        </div>
    </header>

    <!-- MAIN CONTENT -->
    <main class="flex-grow max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8">

        <!-- KPI CARDS METRICS -->
        <section class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
            <!-- Promedio Global -->
            <div class="glass-card rounded-2xl p-5 shadow-sm border-l-4 border-blue-600 hover:shadow-md transition-shadow">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-xs font-bold text-slate-500 uppercase tracking-wider">Promedio Global</p>
                        <h3 id="kpiPromedioGlobal" class="text-3xl font-extrabold text-slate-900 mt-1">241 <span class="text-sm font-normal text-slate-500">pts</span></h3>
                    </div>
                    <div class="w-12 h-12 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center text-xl">
                        <i class="fa-solid fa-chart-line"></i>
                    </div>
                </div>
                <div class="mt-3 space-y-1">
                    <p class="text-xs text-emerald-600 font-bold flex items-center gap-1">
                        <i class="fa-solid fa-arrow-trend-up"></i> +41 pts respecto al año anterior
                    </p>
                    <p class="text-xs text-slate-500 flex items-center gap-1">
                        <span class="text-blue-600 font-semibold"><i class="fa-solid fa-users"></i> 9</span> Estudiantes evaluados
                    </p>
                </div>
            </div>

            <!-- Desviación Estándar -->
            <div class="glass-card rounded-2xl p-5 shadow-sm border-l-4 border-indigo-600 hover:shadow-md transition-shadow">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-xs font-bold text-slate-500 uppercase tracking-wider">Desviación Estándar</p>
                        <h3 id="kpiDesviacion" class="text-3xl font-extrabold text-slate-900 mt-1">32 <span class="text-sm font-normal text-slate-500">pts</span></h3>
                    </div>
                    <div class="w-12 h-12 rounded-xl bg-indigo-50 text-indigo-600 flex items-center justify-center text-xl">
                        <i class="fa-solid fa-arrows-left-right"></i>
                    </div>
                </div>
                <p class="text-xs text-slate-500 mt-3">Dispersión de los puntajes globales</p>
            </div>

            <!-- Mejor Área -->
            <div class="glass-card rounded-2xl p-5 shadow-sm border-l-4 border-emerald-500 hover:shadow-md transition-shadow">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-xs font-bold text-slate-500 uppercase tracking-wider">Área Destacada</p>
                        <h3 id="kpiMejorArea" class="text-2xl font-extrabold text-slate-900 mt-1">Matemáticas</h3>
                    </div>
                    <div class="w-12 h-12 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center text-xl">
                        <i class="fa-solid fa-trophy"></i>
                    </div>
                </div>
                <p class="text-xs text-emerald-600 font-semibold mt-3 flex items-center gap-1">
                    <i class="fa-solid fa-arrow-up"></i> Promedio: <span id="kpiMejorAreaVal">52</span> pts
                </p>
            </div>

            <!-- Área a Reforzar -->
            <div class="glass-card rounded-2xl p-5 shadow-sm border-l-4 border-amber-500 hover:shadow-md transition-shadow">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-xs font-bold text-slate-500 uppercase tracking-wider">Área a Reforzar</p>
                        <h3 id="kpiAreaReforzar" class="text-2xl font-extrabold text-slate-900 mt-1">Inglés</h3>
                    </div>
                    <div class="w-12 h-12 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center text-xl">
                        <i class="fa-solid fa-triangle-exclamation"></i>
                    </div>
                </div>
                <p class="text-xs text-amber-600 font-semibold mt-3 flex items-center gap-1">
                    <i class="fa-solid fa-arrow-down"></i> Promedio: <span id="kpiAreaReforzarVal">44</span> pts
                </p>
            </div>
        </section>

        <!-- CHARTS SECTION -->
        <section class="grid grid-cols-1 lg:grid-cols-12 gap-6">
            <!-- Bar Chart: Promedio por Áreas -->
            <div class="lg:col-span-7 glass-card p-6 rounded-2xl shadow-sm">
                <div class="flex items-center justify-between mb-4">
                    <div>
                        <h2 class="text-lg font-bold text-slate-900">Promedio por Áreas de Conocimiento</h2>
                        <p class="text-xs text-slate-500">Comparativa del rendimiento medio por materia</p>
                    </div>
                    <span class="text-xs bg-slate-100 text-slate-600 px-2.5 py-1 rounded-full font-medium">Escala 0-100</span>
                </div>
                <div class="relative h-72 w-full">
                    <canvas id="barChartAreas"></canvas>
                </div>
            </div>

            <!-- Pie Chart: Distribución por Niveles de Desempeño -->
            <div class="lg:col-span-5 glass-card p-6 rounded-2xl shadow-sm flex flex-col justify-between">
                <div>
                    <div class="flex items-center justify-between mb-4">
                        <div>
                            <h2 class="text-lg font-bold text-slate-900">Niveles de Desempeño</h2>
                            <p class="text-xs text-slate-500">Distribución de estudiantes según rango de Puntaje Global</p>
                        </div>
                    </div>
                    <div class="relative h-60 w-full flex items-center justify-center">
                        <canvas id="pieChartNiveles"></canvas>
                    </div>
                </div>
                <div class="grid grid-cols-3 gap-2 pt-4 border-t border-slate-100 text-center text-xs">
                    <div class="p-2 bg-red-50 rounded-lg text-red-700">
                        <span class="font-bold block text-sm">1</span> Crítico (&lt;200)
                    </div>
                    <div class="p-2 bg-amber-50 rounded-lg text-amber-700">
                        <span class="font-bold block text-sm">3</span> Bajo (200-249)
                    </div>
                    <div class="p-2 bg-orange-50 rounded-lg text-orange-700">
                        <span class="font-bold block text-sm">5</span> Medio (250-299)
                    </div>
                </div>
            </div>
        </section>

        <!-- DIAGNOSTIC MATRIX DAFO -->
        <section>
            <div class="mb-4">
                <h2 class="text-xl font-extrabold text-slate-900 flex items-center gap-2">
                    <i class="fa-solid fa-sliders text-indigo-600"></i> Matriz Diagnóstica DAFO
                </h2>
                <p class="text-xs text-slate-500">Análisis cualitativo y estratégico derivado de los resultados institucionales</p>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
                <!-- Fortalezas -->
                <div class="bg-emerald-50/60 border border-emerald-200 rounded-2xl p-5 shadow-sm">
                    <div class="flex items-center gap-3 mb-3">
                        <div class="w-9 h-9 rounded-xl bg-emerald-600 text-white flex items-center justify-center font-bold text-sm">
                            <i class="fa-solid fa-shield-halved"></i>
                        </div>
                        <h3 class="text-base font-bold text-emerald-900">Fortalezas (Internas)</h3>
                    </div>
                    <ul class="space-y-2 text-xs text-emerald-950">
                        <li class="flex items-start gap-2">
                            <i class="fa-solid fa-circle-check text-emerald-600 mt-0.5"></i>
                            <span><strong>Matemáticas</strong> destaca como la disciplina de mayor desempeño promedio en la institución con <strong>52 pts</strong>.</span>
                        </li>
                        <li class="flex items-start gap-2">
                            <i class="fa-solid fa-circle-check text-emerald-600 mt-0.5"></i>
                            <span><strong>Ciencias Naturales</strong> mantiene un nivel sólido e intermedio de rendimiento con un promedio de <strong>49 pts</strong>.</span>
                        </li>
                        <li class="flex items-start gap-2">
                            <i class="fa-solid fa-circle-check text-emerald-600 mt-0.5"></i>
                            <span>El <strong>56%</strong> de los estudiantes (5 de 9) se ubican en el Nivel Medio de desempeño general.</span>
                        </li>
                    </ul>
                </div>

                <!-- Debilidades -->
                <div class="bg-rose-50/60 border border-rose-200 rounded-2xl p-5 shadow-sm">
                    <div class="flex items-center gap-3 mb-3">
                        <div class="w-9 h-9 rounded-xl bg-rose-600 text-white flex items-center justify-center font-bold text-sm">
                            <i class="fa-solid fa-circle-exclamation"></i>
                        </div>
                        <h3 class="text-base font-bold text-rose-900">Debilidades (Internas)</h3>
                    </div>
                    <ul class="space-y-2 text-xs text-rose-950">
                        <li class="flex items-start gap-2">
                            <i class="fa-solid fa-circle-xmark text-rose-600 mt-0.5"></i>
                            <span><strong>Inglés</strong> representa la mayor oportunidad de mejora con el promedio más bajo (<strong>44 pts</strong>).</span>
                        </li>
                        <li class="flex items-start gap-2">
                            <i class="fa-solid fa-circle-xmark text-rose-600 mt-0.5"></i>
                            <span><strong>Sociales y Ciu.</strong> presenta la mayor dispersión en los puntajes individuales (DE de <strong>10 pts</strong>).</span>
                        </li>
                        <li class="flex items-start gap-2">
                            <i class="fa-solid fa-circle-xmark text-rose-600 mt-0.5"></i>
                            <span>Presencia de estudiantes en Nivel Crítico (&lt;200 pts) que requieren plan de choque prioritario.</span>
                        </li>
                    </ul>
                </div>

                <!-- Oportunidades -->
                <div class="bg-sky-50/60 border border-sky-200 rounded-2xl p-5 shadow-sm">
                    <div class="flex items-center gap-3 mb-3">
                        <div class="w-9 h-9 rounded-xl bg-sky-600 text-white flex items-center justify-center font-bold text-sm">
                            <i class="fa-solid fa-lightbulb"></i>
                        </div>
                        <h3 class="text-base font-bold text-sky-900">Oportunidades (Externas)</h3>
                    </div>
                    <ul class="space-y-2 text-xs text-sky-950">
                        <li class="flex items-start gap-2">
                            <i class="fa-solid fa-arrow-trend-up text-sky-600 mt-0.5"></i>
                            <span>Alto potencial de impulso en el grupo de Nivel Medio para promocionarlos a Niveles Alto y Sobresaliente.</span>
                        </li>
                        <li class="flex items-start gap-2">
                            <i class="fa-solid fa-arrow-trend-up text-sky-600 mt-0.5"></i>
                            <span>Implementación de talleres transversales de Lectura Crítica aplicada a Ciencias Sociales y Ciudadanas.</span>
                        </li>
                        <li class="flex items-start gap-2">
                            <i class="fa-solid fa-arrow-trend-up text-sky-600 mt-0.5"></i>
                            <span>Uso de plataformas digitales interactivas para fortalecimiento focalizado en competencia léxica en Inglés.</span>
                        </li>
                    </ul>
                </div>

                <!-- Amenazas -->
                <div class="bg-amber-50/60 border border-amber-200 rounded-2xl p-5 shadow-sm">
                    <div class="flex items-center gap-3 mb-3">
                        <div class="w-9 h-9 rounded-xl bg-amber-600 text-white flex items-center justify-center font-bold text-sm">
                            <i class="fa-solid fa-triangle-exclamation"></i>
                        </div>
                        <h3 class="text-base font-bold text-amber-900">Amenazas (Externas)</h3>
                    </div>
                    <ul class="space-y-2 text-xs text-amber-950">
                        <li class="flex items-start gap-2">
                            <i class="fa-solid fa-bolt text-amber-600 mt-0.5"></i>
                            <span>Riesgo de estancamiento en la clasificación global de la institución si no se nivela al grupo en rango bajo.</span>
                        </li>
                        <li class="flex items-start gap-2">
                            <i class="fa-solid fa-bolt text-amber-600 mt-0.5"></i>
                            <span>Brecha amplia entre el puntaje máximo (<strong>288 pts</strong>) y el mínimo (<strong>194 pts</strong>).</span>
                        </li>
                        <li class="flex items-start gap-2">
                            <i class="fa-solid fa-bolt text-amber-600 mt-0.5"></i>
                            <span>Aumento de exigencia en estándares nacionales para el acceso a becas y educación superior.</span>
                        </li>
                    </ul>
                </div>
            </div>
        </section>

        <!-- RANKING TABLE SECTION -->
        <section class="glass-card rounded-2xl p-6 shadow-sm">
            <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
                <div>
                    <h2 class="text-xl font-extrabold text-slate-900">Tabla de Posiciones y Resultados Individuales</h2>
                    <p class="text-xs text-slate-500">Listado de estudiantes ordenados de mayor a menor Puntaje Global</p>
                </div>
                
                <div class="flex items-center gap-2">
                    <div class="relative">
                        <i class="fa-solid fa-magnifying-glass absolute left-3 top-2.5 text-slate-400 text-xs"></i>
                        <input type="text" id="searchInput" onkeyup="filterData()" placeholder="Buscar estudiante..." class="pl-8 pr-4 py-1.5 bg-slate-50 text-slate-800 text-xs rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500 w-48 sm:w-64">
                    </div>
                </div>
            </div>

            <!-- Table Container -->
            <div class="overflow-x-auto rounded-xl border border-slate-200">
                <table class="w-full text-left border-collapse" id="studentsTable">
                    <thead>
                        <tr class="bg-slate-100 text-slate-700 text-xs font-bold uppercase tracking-wider border-b border-slate-200">
                            <th class="py-3 px-4 text-center">#</th>
                            <th class="py-3 px-4">Estudiante</th>
                            <th class="py-3 px-4 text-center">L. Crítica</th>
                            <th class="py-3 px-4 text-center">Matemáticas</th>
                            <th class="py-3 px-4 text-center">Sociales</th>
                            <th class="py-3 px-4 text-center">C. Naturales</th>
                            <th class="py-3 px-4 text-center">Inglés</th>
                            <th class="py-3 px-4 text-center">Puntaje Global</th>
                            <th class="py-3 px-4 text-center">Nivel</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody" class="divide-y divide-slate-100 text-xs font-medium text-slate-700">
                        <!-- Dynamic rows populated via JS -->
                    </tbody>
                </table>
            </div>
            
            <div class="mt-4 flex items-center justify-between text-xs text-slate-500">
                <span>Mostrando <strong id="studentCount" class="text-slate-800">9</strong> estudiantes</span>
                <span>Puntajes redondeados sin decimales según regla estricta (&gt;0.5 aproxima superior, &le;0.5 mantiene inferior)</span>
            </div>
        </section>

    </main>

    <!-- FOOTER -->
    <footer class="bg-slate-900 text-slate-400 py-6 border-t border-slate-800 mt-12 text-center text-xs">
        <div class="max-w-7xl mx-auto px-4">
            <p>© 2026 Institución Educativa Simón Bolívar — Sistema Oficial de Reporte de Resultados Saber 11°</p>
        </div>
    </footer>

    <!-- JAVASCRIPT CODE -->
    <script>
        // Custom strict rounding rule function:
        // If decimal > 0.5 round to upper integer, if decimal <= 0.5 keep lower integer.
        function customRound(val) {
            if (val === null || val === undefined || isNaN(val)) return val;
            let floorVal = Math.floor(val);
            let frac = val - floorVal;
            if (frac > 0.5) {
                return Math.ceil(val);
            } else {
                return floorVal;
            }
        }

        // Raw Student Dataset from EXCEL
        const studentsRaw = [
            { num: 11, nombre: "DIAZ GASPAR JULIANA", grado: "11°", lc: 57, mat: 56, soc: 60, nat: 59, ing: 52, pg_raw: 287.692308 },
            { num: 34, nombre: "VELSQUEZ MERCADO LINDA", grado: "11°", lc: 52, mat: 55, soc: 52, nat: 57, ing: 54, pg_raw: 270.000000 },
            { num: 12, nombre: "DIAZ GASPAR ALEJANDO", grado: "11°", lc: 56, mat: 47, soc: 58, nat: 57, ing: 43, pg_raw: 268.076923 },
            { num: 16, nombre: "DE LA CRUZ ROMERO YERALDIN", grado: "11°", lc: 52, mat: 61, soc: 44, nat: 51, ing: 46, pg_raw: 257.692308 },
            { num: 14, nombre: "OQUENDO OVIEDO OSMAIDER", grado: "11°", lc: 47, mat: 60, soc: 48, nat: 46, ing: 51, pg_raw: 251.538462 },
            { num: 13, nombre: "NOVOA RUBIO DANIELA", grado: "11°", lc: 46, mat: 42, soc: 43, nat: 45, ing: 43, pg_raw: 219.615385 },
            { num: 8, nombre: "DORIA JULIO SARA", grado: "11°", lc: 45, mat: 48, soc: 34, nat: 43, ing: 43, pg_raw: 212.692308 },
            { num: 2, nombre: "HOYOS HERNÁNDEZ ESTIVEN", grado: "11°", lc: 38, mat: 52, soc: 38, nat: 44, ing: 36, pg_raw: 212.307692 },
            { num: 15, nombre: "PAYARES CABELLO YULISA", grado: "11°", lc: 35, mat: 49, soc: 34, nat: 40, ing: 30, pg_raw: 193.846154 }
        ];

        // Process data with exact custom rounding
        const students = studentsRaw.map(s => {
            const pgRounded = customRound(s.pg_raw);
            let nivel = "Crítico";
            let badgeClass = "bg-red-100 text-red-800 border-red-200";

            if (pgRounded < 200) {
                nivel = "Crítico";
                badgeClass = "bg-red-100 text-red-800 border-red-200";
            } else if (pgRounded < 250) {
                nivel = "Bajo";
                badgeClass = "bg-amber-100 text-amber-800 border-amber-200";
            } else if (pgRounded < 300) {
                nivel = "Medio";
                badgeClass = "bg-orange-100 text-orange-800 border-orange-200";
            } else if (pgRounded < 400) {
                nivel = "Alto";
                badgeClass = "bg-blue-100 text-blue-800 border-blue-200";
            } else {
                nivel = "Sobresaliente";
                badgeClass = "bg-emerald-100 text-emerald-800 border-emerald-200";
            }

            return {
                ...s,
                pg: pgRounded,
                nivel,
                badgeClass
            };
        });

        // Global Chart Instances
        let barChart = null;
        let pieChart = null;

        // On Page Load
        document.addEventListener("DOMContentLoaded", () => {
            renderDashboard(students);
        });

        function filterData() {
            const selectedGrade = document.getElementById("gradeFilter").value;
            const searchQuery = document.getElementById("searchInput").value.toLowerCase().trim();

            let filtered = students.filter(s => {
                const matchesGrade = (selectedGrade === "ALL" || s.grado === selectedGrade);
                const matchesSearch = s.nombre.toLowerCase().includes(searchQuery) || s.num.toString().includes(searchQuery);
                return matchesGrade && matchesSearch;
            });

            // Always keep sorted High to Low
            filtered.sort((a, b) => b.pg - a.pg);

            renderDashboard(filtered);
        }

        function renderDashboard(dataList) {
            renderTable(dataList);
            updateKPIs(dataList);
            updateCharts(dataList);
        }

        function renderTable(dataList) {
            const tbody = document.getElementById("tableBody");
            tbody.innerHTML = "";
            document.getElementById("studentCount").innerText = dataList.length;

            if (dataList.length === 0) {
                tbody.innerHTML = `
                    <tr>
                        <td colspan="9" class="py-8 text-center text-slate-400">
                            <i class="fa-solid fa-folder-open text-2xl mb-2 block"></i>
                            No se encontraron estudiantes con los criterios seleccionados.
                        </td>
                    </tr>`;
                return;
            }

            dataList.forEach((s, index) => {
                const tr = document.createElement("tr");
                tr.className = "hover:bg-slate-50/80 transition-colors border-b border-slate-100";
                
                tr.innerHTML = `
                    <td class="py-3 px-4 text-center font-bold text-slate-400">${index + 1}</td>
                    <td class="py-3 px-4 font-semibold text-slate-900">${s.nombre}</td>
                    <td class="py-3 px-4 text-center font-medium">${s.lc}</td>
                    <td class="py-3 px-4 text-center font-medium">${s.mat}</td>
                    <td class="py-3 px-4 text-center font-medium">${s.soc}</td>
                    <td class="py-3 px-4 text-center font-medium">${s.nat}</td>
                    <td class="py-3 px-4 text-center font-medium">${s.ing}</td>
                    <td class="py-3 px-4 text-center font-extrabold text-slate-900 text-sm bg-slate-50/50">${s.pg}</td>
                    <td class="py-3 px-4 text-center">
                        <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-bold border ${s.badgeClass}">
                            ${s.nivel}
                        </span>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function updateKPIs(dataList) {
            if (dataList.length === 0) {
                document.getElementById("kpiPromedioGlobal").innerHTML = `0 <span class="text-sm font-normal text-slate-500">pts</span>`;
                document.getElementById("kpiDesviacion").innerHTML = `0 <span class="text-sm font-normal text-slate-500">pts</span>`;
                return;
            }

            // Averages & STDev calculation based on raw values then rounded
            const pgSum = dataList.reduce((acc, curr) => acc + curr.pg_raw, 0);
            const pgMeanRaw = pgSum / dataList.length;
            const pgMeanRounded = customRound(pgMeanRaw);

            // Standard Deviation
            let stdDevRounded = 0;
            if (dataList.length > 1) {
                const variance = dataList.reduce((acc, curr) => acc + Math.pow(curr.pg_raw - pgMeanRaw, 2), 0) / (dataList.length - 1);
                const stdDevRaw = Math.sqrt(variance);
                stdDevRounded = customRound(stdDevRaw);
            }

            // Areas calculation
            const areaMeansRaw = {
                "Lectura Crítica": dataList.reduce((a, c) => a + c.lc, 0) / dataList.length,
                "Matemáticas": dataList.reduce((a, c) => a + c.mat, 0) / dataList.length,
                "Sociales": dataList.reduce((a, c) => a + c.soc, 0) / dataList.length,
                "Ciencias Naturales": dataList.reduce((a, c) => a + c.nat, 0) / dataList.length,
                "Inglés": dataList.reduce((a, c) => a + c.ing, 0) / dataList.length,
            };

            const areaMeans = {};
            for (let k in areaMeansRaw) {
                areaMeans[k] = customRound(areaMeansRaw[k]);
            }

            let bestArea = "Matemáticas";
            let bestVal = -1;
            let worstArea = "Inglés";
            let worstVal = 999;

            for (let area in areaMeans) {
                if (areaMeans[area] > bestVal) {
                    bestVal = areaMeans[area];
                    bestArea = area;
                }
                if (areaMeans[area] < worstVal) {
                    worstVal = areaMeans[area];
                    worstArea = area;
                }
            }

            document.getElementById("kpiPromedioGlobal").innerHTML = `${pgMeanRounded} <span class="text-sm font-normal text-slate-500">pts</span>`;
            document.getElementById("kpiDesviacion").innerHTML = `${stdDevRounded} <span class="text-sm font-normal text-slate-500">pts</span>`;
            document.getElementById("kpiMejorArea").innerText = bestArea;
            document.getElementById("kpiMejorAreaVal").innerText = bestVal;
            document.getElementById("kpiAreaReforzar").innerText = worstArea;
            document.getElementById("kpiAreaReforzarVal").innerText = worstVal;
        }

        function updateCharts(dataList) {
            if (dataList.length === 0) return;

            // Compute Area Means
            const areaMeans = {
                "Lectura Crítica": customRound(dataList.reduce((a, c) => a + c.lc, 0) / dataList.length),
                "Matemáticas": customRound(dataList.reduce((a, c) => a + c.mat, 0) / dataList.length),
                "Sociales": customRound(dataList.reduce((a, c) => a + c.soc, 0) / dataList.length),
                "Ciencias Naturales": customRound(dataList.reduce((a, c) => a + c.nat, 0) / dataList.length),
                "Inglés": customRound(dataList.reduce((a, c) => a + c.ing, 0) / dataList.length),
            };

            // Bar Chart Initialization/Update
            const barCtx = document.getElementById("barChartAreas").getContext("2d");
            if (barChart) barChart.destroy();

            barChart = new Chart(barCtx, {
                type: 'bar',
                data: {
                    labels: Object.keys(areaMeans),
                    datasets: [{
                        label: 'Puntaje Promedio',
                        data: Object.values(areaMeans),
                        // Distinct colors per bar
                        backgroundColor: [
                            '#6366f1', // Indigo for Lectura Crítica
                            '#0284c7', // Sky Blue for Matemáticas
                            '#e11d48', // Rose for Sociales
                            '#059669', // Emerald for Ciencias Naturales
                            '#d97706'  // Amber for Inglés
                        ],
                        borderRadius: 8,
                        borderSkipped: false
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    return ` Promedio: ${context.parsed.y} pts`;
                                }
                            }
                        }
                    },
                    scales: {
                        y: {
                            beginAtZero: true,
                            max: 100,
                            ticks: {
                                stepSize: 20
                            },
                            grid: {
                                color: '#f1f5f9'
                            }
                        },
                        x: {
                            grid: { display: false }
                        }
                    }
                }
            });

            // Performance Levels Counting
            const levelsCount = {
                "Crítico (<200)": 0,
                "Bajo (200-249)": 0,
                "Medio (250-299)": 0,
                "Alto (300-399)": 0,
                "Sobresaliente (≥400)": 0
            };

            dataList.forEach(s => {
                if (s.pg < 200) levelsCount["Crítico (<200)"]++;
                else if (s.pg < 250) levelsCount["Bajo (200-249)"]++;
                else if (s.pg < 300) levelsCount["Medio (250-299)"]++;
                else if (s.pg < 400) levelsCount["Alto (300-399)"]++;
                else levelsCount["Sobresaliente (≥400)"]++;
            });

            // Pie Chart Initialization/Update
            const pieCtx = document.getElementById("pieChartNiveles").getContext("2d");
            if (pieChart) pieChart.destroy();

            pieChart = new Chart(pieCtx, {
                type: 'doughnut',
                data: {
                    labels: Object.keys(levelsCount),
                    datasets: [{
                        data: Object.values(levelsCount),
                        backgroundColor: [
                            '#ef4444', // Red - Crítico
                            '#eab308', // Yellow - Bajo
                            '#f97316', // Orange - Medio
                            '#3b82f6', // Blue - Alto
                            '#22c55e'  // Green - Sobresaliente
                        ],
                        borderWidth: 2,
                        borderColor: '#ffffff'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                boxWidth: 12,
                                font: { size: 11 }
                            }
                        },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    const count = context.parsed;
                                    const total = dataList.length;
                                    const percentage = total > 0 ? customRound((count / total) * 100) : 0;
                                    return ` ${context.label}: ${count} estudiantes (${percentage}%)`;
                                }
                            }
                        }
                    },
                    cutout: '60%'
                }
            });
        }

        // Export Filtered Table to Excel using SheetJS
        function exportToExcel() {
            const selectedGrade = document.getElementById("gradeFilter").value;
            const searchQuery = document.getElementById("searchInput").value.toLowerCase().trim();

            let currentFiltered = students.filter(s => {
                const matchesGrade = (selectedGrade === "ALL" || s.grado === selectedGrade);
                const matchesSearch = s.nombre.toLowerCase().includes(searchQuery) || s.num.toString().includes(searchQuery);
                return matchesGrade && matchesSearch;
            });

            // Ensure High to Low sorting
            currentFiltered.sort((a, b) => b.pg - a.pg);

            // Build dataset for sheet
            const sheetData = currentFiltered.map((s, index) => ({
                "Posición": index + 1,
                "Nombre y Apellidos": s.nombre,
                "Grado": s.grado,
                "Lectura Crítica": s.lc,
                "Matemáticas": s.mat,
                "Sociales y Competencias Ciudadanas": s.soc,
                "Ciencias Naturales": s.nat,
                "Inglés": s.ing,
                "Puntaje Global": s.pg,
                "Nivel de Desempeño": s.nivel
            }));

            const worksheet = XLSX.utils.json_to_sheet(sheetData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Resultados Saber 11");

            // Auto-adjust column widths
            const max_widths = [
                { wch: 10 }, { wch: 32 }, { wch: 10 },
                { wch: 16 }, { wch: 14 }, { wch: 34 }, { wch: 18 },
                { wch: 10 }, { wch: 16 }, { wch: 20 }
            ];
            worksheet['!cols'] = max_widths;

            // Save file
            XLSX.writeFile(workbook, "Resultados_Saber11_IE_Simon_Bolivar.xlsx");
        }
    </script>
</body>
</html>
