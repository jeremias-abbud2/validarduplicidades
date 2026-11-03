import { useState, useCallback, useRef, DragEvent, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  AlertCircle, 
  CheckCircle2, 
  Clock, 
  User, 
  Building2, 
  FileText,
  Trash2,
  FileDown,
  Search,
  Filter,
  X,
  BarChart3,
  TrendingUp,
  Calendar,
  AlertTriangle,
  PieChart
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from './lib/utils';

// Types for the report data
interface TransactionRow {
  'DATA DA TRAN'?: string | number | Date;
  'TIPO DE TRANSAÇÃO'?: string;
  'BENEFICIÁRIO'?: string;
  'MATRÍCULA'?: string | number;
  'PRESTADOR'?: string;
  'PROCEDIMENTO'?: string;
  'VALOR'?: string | number;
  'QUANT'?: string | number;
  'SITUAÇÃO'?: string;
  'Nº CONSELHO'?: string | number;
  [key: string]: any;
}

interface DuplicateGroup {
  fingerprint: string;
  rows: TransactionRow[];
  params: {
    prestador: string;
    prestador_matricula: string;
    beneficiario: string;
    beneficiario_matricula: string;
    procedimento: string;
    especialidade: string;
    situacao: string;
    tipo: string;
    quantidade: string;
    conselho: string;
    data: string;
    conselhoExecTipo: string;
    conselhoExecNum: string;
    conselhoExecUf: string;
  };
}

interface FilterState {
  prestador: string;
  beneficiario: string;
  matricula: string;
  procedimento: string;
  data: string;
  situacao: string;
  intervalCategory: string;
  conselhoExecutante: string;
}

export default function App() {
  const [data, setData] = useState<TransactionRow[] | null>(null);
  const [duplicateGroups, setDuplicateGroups] = useState<DuplicateGroup[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [transactionCount, setTransactionCount] = useState<number>(0);
  const [filters, setFilters] = useState<FilterState>({
    prestador: '',
    beneficiario: '',
    matricula: '',
    procedimento: '',
    data: '',
    situacao: '',
    intervalCategory: '',
    conselhoExecutante: ''
  });

  const [uniqueValues, setUniqueValues] = useState<{
    prestadores: string[];
    procedimentos: string[];
    situacoes: string[];
    datas: string[];
    conselhoExecutantes: string[];
  }>({
    prestadores: [],
    procedimentos: [],
    situacoes: [],
    datas: [],
    conselhoExecutantes: [],
  });
  
  const [activeTab, setActiveTab] = useState<'table' | 'dashboard'>('table');

  const categoryStats = useMemo(() => {
    let high = 0;
    let medium = 0;
    let low = 0;
    duplicateGroups.forEach(g => {
      const info = getGroupTimeInterval(g.rows);
      if (info.level === 'high') high++;
      else if (info.level === 'medium') medium++;
      else low++;
    });
    return { high, medium, low, total: duplicateGroups.length };
  }, [duplicateGroups]);

  const dashboardStats = useMemo(() => {
    const daysMap: Record<string, number> = {};
    const hoursMap: Record<string, number> = {};
    const providersMap: Record<string, number> = {};
    const proceduresMap: Record<string, number> = {};

    duplicateGroups.forEach(group => {
      const d = group.params.data || 'Data não informada';
      daysMap[d] = (daysMap[d] || 0) + 1;

      const prov = group.params.prestador || 'Não informado';
      providersMap[prov] = (providersMap[prov] || 0) + 1;

      const proc = group.params.procedimento || 'Não informado';
      proceduresMap[proc] = (proceduresMap[proc] || 0) + 1;

      group.rows.forEach(row => {
        const dateVal = getColValue(row, 'DATA');
        const hourVal = getColValue(row, 'HORA');
        let { time } = normalizeDate(dateVal);
        if ((!time || time === '00:00') && hourVal !== undefined && hourVal !== null && hourVal !== '') {
          const hNorm = normalizeDate(hourVal);
          if (hNorm.time) time = hNorm.time;
          else if (typeof hourVal === 'string' && hourVal.includes(':')) time = hourVal.trim().substring(0, 5);
        }
        if (time && time !== '00:00' && time.includes(':')) {
          const hourGroup = time.split(':')[0] + ':00';
          hoursMap[hourGroup] = (hoursMap[hourGroup] || 0) + 1;
        }
      });
    });

    const topDays = Object.entries(daysMap)
      .map(([day, count]) => ({ day, count }))
      .sort((a, b) => b.count - a.count);

    const topHours = Object.entries(hoursMap)
      .map(([hour, count]) => ({ hour, count }))
      .sort((a, b) => b.count - a.count);

    const topProviders = Object.entries(providersMap)
      .map(([provider, count]) => ({ provider, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    const topProcedures = Object.entries(proceduresMap)
      .map(([procedure, count]) => ({ procedure, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    return { topDays, topHours, topProviders, topProcedures };
  }, [duplicateGroups]);

  const exportDashboardReport = () => {
    if (duplicateGroups.length === 0) return;

    const wb = XLSX.utils.book_new();

    const summaryData = [
      { Métrica: 'Total de Registros Analisados', Valor: transactionCount },
      { Métrica: 'Total de Grupos de Duplicidade', Valor: duplicateGroups.length },
      { Métrica: 'Duplicidades Bem Prováveis (≤ 30min)', Valor: categoryStats.high },
      { Métrica: 'Duplicidades Pouco Menos Prováveis (30m - 1h)', Valor: categoryStats.medium },
      { Métrica: 'Outras Duplicidades (> 1h)', Valor: categoryStats.low },
    ];
    const wsSummary = XLSX.utils.json_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, wsSummary, "Resumo Executivo");

    const wsDays = XLSX.utils.json_to_sheet(dashboardStats.topDays.map(d => ({ 'Data': d.day, 'Quantidade de Duplicidades': d.count })));
    XLSX.utils.book_append_sheet(wb, wsDays, "Picos por Dia");

    const wsHours = XLSX.utils.json_to_sheet(dashboardStats.topHours.map(h => ({ 'Faixa Horária': h.hour, 'Ocorrências': h.count })));
    XLSX.utils.book_append_sheet(wb, wsHours, "Picos por Horário");

    const wsProv = XLSX.utils.json_to_sheet(dashboardStats.topProviders.map(p => ({ 'Prestador': p.provider, 'Duplicidades': p.count })));
    XLSX.utils.book_append_sheet(wb, wsProv, "Top Prestadores");

    XLSX.writeFile(wb, "Dashboard_Auditoria_Duplicidades.xlsx");
  };

  const fileInputRef = useRef<HTMLInputElement>(null);

  function normalizeDate(val: any): { date: string; time: string } {
    if (val === undefined || val === null || val === '') return { date: '', time: '' };
    
    try {
      if (typeof val === 'number') {
        // Caso Excel: pode ser data+hora ou apenas hora (valor < 1)
        const date = XLSX.SSF.parse_date_code(val);
        const d = val >= 1 ? `${date.d.toString().padStart(2, '0')}/${date.m.toString().padStart(2, '0')}/${date.y}` : '';
        const t = `${date.H.toString().padStart(2, '0')}:${date.M.toString().padStart(2, '0')}`;
        return { date: d, time: t };
      }

      if (val instanceof Date) {
        if (isNaN(val.getTime())) return { date: '', time: '' };
        const d = `${val.getDate().toString().padStart(2, '0')}/${(val.getMonth() + 1).toString().padStart(2, '0')}/${val.getFullYear()}`;
        const t = `${val.getHours().toString().padStart(2, '0')}:${val.getMinutes().toString().padStart(2, '0')}`;
        return { date: d, time: t };
      }

      const str = String(val).trim();
      if (!str) return { date: '', time: '' };

      const parts = str.split(/[\sT]/);
      
      // Caso string contenha apenas horário (ex: "12:30" ou "12:30:00")
      if (parts.length === 1 && parts[0].includes(':')) {
        return { date: '', time: parts[0].substring(0, 5) };
      }

      const datePart = parts[0] || '';
      const timePart = (parts[1] || '').substring(0, 5); 
      
      return {
        date: datePart,
        time: timePart
      };
    } catch (e) {
      const s = String(val);
      return s.includes(':') ? { date: '', time: s.substring(0, 5) } : { date: s, time: '' };
    }
  };

  function getGroupTimeInterval(rows: TransactionRow[]) {
    const times: number[] = [];
    rows.forEach(row => {
      const dateVal = getColValue(row, 'DATA');
      const hourVal = getColValue(row, 'HORA');
      let { time } = normalizeDate(dateVal);
      if ((!time || time === '00:00') && hourVal !== undefined && hourVal !== null && hourVal !== '') {
        const hNorm = normalizeDate(hourVal);
        if (hNorm.time) {
          time = hNorm.time;
        } else if (typeof hourVal === 'string' && hourVal.includes(':')) {
          time = hourVal.trim().substring(0, 5);
        }
      }
      if (time && time !== '00:00' && time !== '--:--' && time.includes(':')) {
        const [hh, mm] = time.split(':').map(Number);
        if (!isNaN(hh) && !isNaN(mm)) {
          times.push(hh * 60 + mm);
        }
      }
    });

    if (times.length < 2) return { minDiff: null, label: '', level: 'normal' };

    times.sort((a, b) => a - b);
    let minDiff = Infinity;
    for (let i = 1; i < times.length; i++) {
      const diff = times[i] - times[i - 1];
      if (diff < minDiff) minDiff = diff;
    }

    if (minDiff <= 30) {
      return { minDiff, label: `≤ 30m (Bem Provável)`, level: 'high' };
    } else if (minDiff <= 60) {
      return { minDiff, label: `30m - 1h (Pouco Menos)`, level: 'medium' };
    } else {
      return { minDiff, label: `> 1h`, level: 'low' };
    }
  };

  function getColValue(row: any, concept: string): any {
    const keys = Object.keys(row);
    
    // Normalização agressiva para lidar com diferentes encodings e caracteres especiais
    const normalize = (s: string) => {
      if (!s) return '';
      return String(s)
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '') // Remove acentos
        .toUpperCase()
        .replace(/\s+/g, '') // Remove espaços
        .replace(/[^A-Z0-9]/g, ''); // Remove tudo que não é alfanumérico
    };

    const aliases: Record<string, string[]> = {
      'DATA': ['DATA_TRAN', 'DATA DA TRAN', 'DT_TRAN', 'DATA', 'DIA', 'DATA_TRANSICAO', 'DATA_ATENDIMENTO', 'DATA DA TRANSACAO'],
      'HORA': ['HORARIO_', 'HORA', 'HORARIO', 'TIME', 'HORA_TRAN', 'HORA_ATENDIMENTO'],
      'BENE_NOME': ['NOME_BENEFICIARIO', 'BENEFICIÁRIO', 'NOME_BE', 'NM_BENEFICIARIO', 'PACIENTE', 'NOME', 'BENEFICIARIO', 'NOME DO BENEFICIARIO'],
      'BENE_MATRICULA': ['MATRICULA_BENEFICIARIO', 'MATRICULA_BENE', 'MATRÍCULA', 'MATRICUL', 'MATRICULA', 'COD_BENE', 'CARTEIRA', 'MATRICULA_BENEFICIARIO_DV', 'MATRICULA_PACIENTE', 'MATRICULA DO BENEFICIARIO'],
      'PREST_NOME': ['NOME_PRESTADOR', 'NM_PRESTADOR', 'PRESTADOR', 'NOME_PR', 'UNIDADE', 'NOME_PRESTADOR_EXECUTANTE', 'NOME DO PRESTADOR'],
      'PREST_MATRICULA': ['MATRICULA_PRESTADOR', 'COD_PRESTADOR', 'CNPJ_PRESTADOR', 'MATRICULA DO PRESTADOR'],
      'PROCEDIMENTO': ['PROCEDIMI', 'PROCEDIMENTO', 'COD_PROC', 'PROCED', 'CODIGO', 'PROCE', 'ITEM', 'PROC'],
      'QUANTIDADE': ['QUANTIDADE', 'QUANT', 'QTD', 'QUANT_SOLIC', 'QTDE'],
      'SITUACAO': ['SITUAÇÃO', 'SITUACAO', 'STATUS', 'RESULTADO', 'SIT'],
      'CONSELHO': ['Nº CONSELHO', 'NUMERO_CONSELHO', 'CONSELHO', 'CRM', 'NUMERO_', 'REGISTRO_CONSELHO', 'CONSELHO PROFISSIONAL'],
      'TIPO': ['TIPO_TRA', 'TIPO DE TRANSAÇÃO', 'TIPO', 'TIPO_TRANSACAO', 'TIPO DE TRANSACAO', 'TIPO DE SOLICITACAO', 'TIPO_SOLICITACAO'],
      'ESPECIALIDADE': ['ESPECIALIDADE', 'ESPEC', 'AREA_MEDICA', 'NOME_ESPECIALIDADE', 'ESPECIALIDADE_PRESTADOR'],
      'TIPO_CONSELHO_EXECUTANTE': ['TIPO_CONSELHO_EXECUTANTE', 'TIPO_CONSELHO_EXEC', 'CONSELHO_EXECUTANTE_TIPO', 'TIPO_CONSELHO', 'CONSELHO_TIPO'],
      'NUMERO_CONSELHO_EXECUTANTE': ['NUMERO_CONSELHO_EXECUTANTE', 'NUM_CONSELHO_EXECUTANTE', 'NUMERO_CONSELHO_EXEC', 'CRM_EXECUTANTE', 'NUM_CONSELHO_EXEC'],
      'UF_CONSELHO_EXECUTANTE': ['UF_CONSELHO_EXECUTANTE', 'UF_CONSELHO_EXEC', 'CONSELHO_EXECUTANTE_UF', 'UF_CRM_EXECUTANTE']
    };

    const targets = aliases[concept] || [concept];
    const normalizedTargets = targets.map(normalize);

    // 1. Prioridade Máxima: Match Exato (sem espaços e case-insensitive)
    for (const target of targets) {
      const found = keys.find(k => k.trim().toUpperCase() === target.toUpperCase());
      if (found) return row[found];
    }

    // 2. Prioridade Média: Match Normalizado
    const candidates = keys.map(k => {
      const normalizedK = normalize(k);
      let bestScore = 0;
      
      normalizedTargets.forEach(t => {
        if (normalizedK === t) bestScore = 100;
        else if (normalizedK.startsWith(t)) bestScore = Math.max(bestScore, 80);
        else if (normalizedK.includes(t)) bestScore = Math.max(bestScore, 50);
      });

      // Penalização para evitar confusão entre Beneficiário e Prestador
      const upperK = k.toUpperCase();
      if ((concept === 'BENE_NOME' || concept === 'BENE_MATRICULA') && upperK.includes('PREST')) bestScore -= 60;
      if ((concept === 'PREST_NOME' || concept === 'PREST_MATRICULA') && upperK.includes('BENE')) bestScore -= 60;

      return { key: k, score: bestScore };
    }).filter(c => c.score > 40);

    candidates.sort((a, b) => b.score - a.score);
    
    return candidates.length > 0 ? row[candidates[0].key] : undefined;
  };

  const processFile = async (file: File) => {
    setIsProcessing(true);
    setFileName(file.name);
    setError(null);
    setData(null);
    setDuplicateGroups([]);

    try {
      let jsonData: TransactionRow[] = [];
      
      // 1. Tentar leitura binária primeiro (funciona para XLSX, XLS e CSVs em UTF-8)
      const arrayBuffer = await file.arrayBuffer();
      const binaryData = new Uint8Array(arrayBuffer);
      
      try {
        const workbook = XLSX.read(binaryData, { 
          type: 'array', 
          cellDates: true,
          dateNF: 'dd/mm/yyyy hh:mm:ss',
          cellText: false,
          cellNF: false,
          cellStyles: false
        });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const initialData = XLSX.utils.sheet_to_json(worksheet, { defval: '' }) as TransactionRow[];
        
        // Se conseguimos ler e parece ter colunas suficientes, usamos isso
        if (initialData.length > 0 && Object.keys(initialData[0]).length > 2) {
          jsonData = initialData;
        }
      } catch (readErr) {
        console.warn("Binary read failed or insufficient columns, trying text-based fallbacks...", readErr);
      }

      // 2. Se não funcionou ou parece ser um CSV com delimitador/encoding especial
      if (jsonData.length === 0) {
        const encodings = ['Windows-1252', 'ISO-8859-1', 'UTF-8'];
        let maxCols = 0;
        let bestData: TransactionRow[] = [];

        for (const encoding of encodings) {
          try {
            const text = await new Promise<string>((resolve, reject) => {
              const reader = new FileReader();
              reader.onload = () => resolve(reader.result as string);
              reader.onerror = reject;
              reader.readAsText(file, encoding);
            });

            // Pula arquivos que parecem ser binários (zips/xlsx) lidos como texto
            if (text.startsWith('PK\x03\x04') || text.includes('\x00')) {
              continue;
            }

            const firstLine = text.split('\n')[0];
            const hasSemicolon = firstLine.includes(';');
            
            // Tenta ler com SheetJS mas tratando como string
            const workbook = XLSX.read(text, { 
              type: 'string', 
              FS: hasSemicolon ? ';' : ',',
              cellDates: true 
            });
            
            const firstSheetName = workbook.SheetNames[0];
            const currentData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { defval: '' }) as TransactionRow[];
            
            if (currentData.length > 0) {
              const cols = Object.keys(currentData[0]).length;
              if (cols > maxCols) {
                maxCols = cols;
                bestData = currentData;
                if (cols > 5) break; 
              }
            }
          } catch (e) {
            console.warn(`Failed reading text with ${encoding}`, e);
          }
        }
        jsonData = bestData;
      }

      if (jsonData.length === 0) {
        throw new Error('O arquivo está vazio ou não pôde ser interpretado.');
      }

      setData(jsonData);
      setTransactionCount(jsonData.length);
      analyzeDuplicates(jsonData);
    } catch (err: any) {
      setError(err.message || 'Erro ao processar o arquivo.');
    } finally {
      setIsProcessing(false);
    }
  };

  const analyzeDuplicates = (rows: TransactionRow[]) => {
    const groupsMap: Record<string, TransactionRow[]> = {};
    const prestadoresSet = new Set<string>();
    const procedimentosSet = new Set<string>();
    const situacoesSet = new Set<string>();
    const datasSet = new Set<string>();
    const conselhoExecSet = new Set<string>();
    
    const normalizeText = (v: any) => String(v ?? '').trim().toLowerCase();
    const normalizeRaw = (v: any) => String(v ?? '').trim();
    const normalizeNum = (v: any) => {
      if (v === undefined || v === null || v === '') return '';
      const n = parseFloat(String(v));
      return isNaN(n) ? normalizeText(v) : String(n);
    };

    const robustNormalize = (s: any) => {
      if (s === undefined || s === null || s === '') return '';
      return String(s)
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '') // Remove acentos
        .toUpperCase()
        .replace(/\s+/g, '') // Remove espaços
        .replace(/[^A-Z0-9]/g, ''); // Remove tudo que não é alfanumérico
    };

    rows.forEach(row => {
      const keys = Object.keys(row);
      const rawDateVal = getColValue(row, 'DATA');
      const { date } = normalizeDate(rawDateVal);

      const prestadorNome = normalizeRaw(getColValue(row, 'PREST_NOME'));
      const prestadorMatricula = robustNormalize(getColValue(row, 'PREST_MATRICULA'));
      const beneficiarioNome = normalizeRaw(getColValue(row, 'BENE_NOME'));
      const beneficiarioMatriculaOriginal = robustNormalize(getColValue(row, 'BENE_MATRICULA'));
      
      let beneficiarioMatricula = beneficiarioMatriculaOriginal;
      // Secondary check for DV if it wasn't already part of the matricula - stricter check
      const dvCol = keys.find(k => {
        const uk = k.toUpperCase();
        return (uk === 'DV' || uk.includes('_DV') || uk.includes(' DV')) && 
               (uk.includes('MATRICULA') || uk.includes('BENE') || uk.includes('PACIENTE'));
      });

      if (dvCol && row[dvCol]) {
        const dv = String(row[dvCol]).trim();
        if (dv && !beneficiarioMatricula.endsWith(dv)) {
          beneficiarioMatricula = `${beneficiarioMatricula}${dv}`;
        }
      }

      const procedimento = normalizeRaw(getColValue(row, 'PROCEDIMENTO'));
      const especialidade = normalizeRaw(getColValue(row, 'ESPECIALIDADE'));
      const situacao = normalizeRaw(getColValue(row, 'SITUACAO'));
      const tipo = normalizeRaw(getColValue(row, 'TIPO'));
      const quantidade = normalizeNum(getColValue(row, 'QUANTIDADE'));
      const conselho = robustNormalize(getColValue(row, 'CONSELHO'));
      
      const cExecTipo = robustNormalize(getColValue(row, 'TIPO_CONSELHO_EXECUTANTE'));
      const cExecNum = robustNormalize(getColValue(row, 'NUMERO_CONSELHO_EXECUTANTE'));
      const cExecUf = robustNormalize(getColValue(row, 'UF_CONSELHO_EXECUTANTE'));

      const params = [
        robustNormalize(prestadorNome),
        prestadorMatricula,
        date,
        robustNormalize(beneficiarioNome),
        beneficiarioMatricula,
        robustNormalize(procedimento),
        robustNormalize(especialidade),
        robustNormalize(situacao),
        robustNormalize(tipo),
        quantidade,
        conselho,
        cExecTipo,
        cExecNum,
        cExecUf,
      ];

      const fingerprint = params.join('|');
      
      if (!groupsMap[fingerprint]) {
        groupsMap[fingerprint] = [];
      }
      groupsMap[fingerprint].push(row);
    });

    const calculatedGroups: DuplicateGroup[] = Object.entries(groupsMap)
      .filter(([_, groupRows]) => groupRows.length > 1)
      .map(([fingerprint, groupRows]) => {
        const first = groupRows[0];
        const { date } = normalizeDate(getColValue(first, 'DATA'));
        
        const bMat = normalizeText(getColValue(first, 'BENE_MATRICULA'));
        const pMat = normalizeText(getColValue(first, 'PREST_MATRICULA'));

        const prestador = String(getColValue(first, 'PREST_NOME') ?? '');
        const procedimento = String(getColValue(first, 'PROCEDIMENTO') ?? '');
        const situacao = String(getColValue(first, 'SITUACAO') ?? '');

        const cExecTipoVal = String(getColValue(first, 'TIPO_CONSELHO_EXECUTANTE') ?? '');
        const cExecNumVal = String(getColValue(first, 'NUMERO_CONSELHO_EXECUTANTE') ?? '');
        const cExecUfVal = String(getColValue(first, 'UF_CONSELHO_EXECUTANTE') ?? '');
        const cExecStr = [cExecTipoVal, cExecNumVal, cExecUfVal].filter(Boolean).join(' ');
        if (cExecStr) conselhoExecSet.add(cExecStr);

        if (prestador) prestadoresSet.add(prestador);
        if (procedimento) procedimentosSet.add(procedimento);
        if (situacao) situacoesSet.add(situacao);
        if (date) datasSet.add(date);

        return {
          fingerprint,
          rows: groupRows,
          params: {
            prestador,
            prestador_matricula: pMat,
            beneficiario: String(getColValue(first, 'BENE_NOME') ?? ''),
            beneficiario_matricula: bMat,
            procedimento,
            especialidade: String(getColValue(first, 'ESPECIALIDADE') ?? ''),
            situacao,
            tipo: String(getColValue(first, 'TIPO') ?? ''),
            quantidade: String(getColValue(first, 'QUANTIDADE') ?? ''),
            conselho: String(getColValue(first, 'CONSELHO') ?? ''),
            data: date,
            conselhoExecTipo: cExecTipoVal,
            conselhoExecNum: cExecNumVal,
            conselhoExecUf: cExecUfVal
          }
        };
      });

    setUniqueValues({
      prestadores: Array.from(prestadoresSet).sort(),
      procedimentos: Array.from(procedimentosSet).sort(),
      situacoes: Array.from(situacoesSet).sort(),
      datas: Array.from(datasSet).sort((a, b) => {
        const [da, ma, ya] = a.split('/').map(Number);
        const [db, mb, yb] = b.split('/').map(Number);
        return new Date(ya, ma - 1, da).getTime() - new Date(yb, mb - 1, db).getTime();
      }),
      conselhoExecutantes: Array.from(conselhoExecSet).sort(),
    });

    setDuplicateGroups(calculatedGroups);
  };

  const filteredDuplicates = useMemo(() => {
    return duplicateGroups.filter(group => {
      const matchPrestador = !filters.prestador || group.params.prestador.toLowerCase().includes(filters.prestador.toLowerCase());
      const matchBeneficiario = !filters.beneficiario || group.params.beneficiario.toLowerCase().includes(filters.beneficiario.toLowerCase());
      const matchMatricula = !filters.matricula || 
        group.params.beneficiario_matricula.toLowerCase().includes(filters.matricula.toLowerCase()) ||
        group.params.prestador_matricula.toLowerCase().includes(filters.matricula.toLowerCase());
      const matchProcedimento = !filters.procedimento || group.params.procedimento.toLowerCase().includes(filters.procedimento.toLowerCase());
      const matchData = !filters.data || group.params.data.includes(filters.data);
      const matchSituacao = !filters.situacao || group.params.situacao.toLowerCase().includes(filters.situacao.toLowerCase());
      
      const cExecFull = `${group.params.conselhoExecTipo} ${group.params.conselhoExecNum} ${group.params.conselhoExecUf}`.toLowerCase();
      const matchConselhoExec = !filters.conselhoExecutante || 
        cExecFull.includes(filters.conselhoExecutante.toLowerCase()) ||
        group.params.conselhoExecNum.toLowerCase().includes(filters.conselhoExecutante.toLowerCase());

      const intervalInfo = getGroupTimeInterval(group.rows);
      const matchInterval = !filters.intervalCategory || intervalInfo.level === filters.intervalCategory;

      return matchPrestador && matchBeneficiario && matchMatricula && matchProcedimento && matchData && matchSituacao && matchConselhoExec && matchInterval;
    });
  }, [duplicateGroups, filters]);

  const handleDrop = (e: DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || file.name.endsWith('.csv'))) {
      processFile(file);
    } else {
      setError('Por favor, envie um arquivo Excel (.xlsx, .xls) ou CSV.');
    }
  };

  const exportResults = () => {
    if (filteredDuplicates.length === 0) return;

    const exportData = filteredDuplicates.flatMap(group => {
      const intervalInfo = getGroupTimeInterval(group.rows);
      return group.rows.map(row => {
        const { date, time } = normalizeDate(getColValue(row, 'DATA'));
        return {
          ...row,
          'DATA_NORMALIZADA': date,
          'HORARIO': time,
          'ID_BENEFICIARIO': getColValue(row, 'BENE_MATRICULA'),
          'ID_PRESTADOR': getColValue(row, 'PREST_MATRICULA'),
          'CATEGORIA_INTERVALO': intervalInfo.label || 'Outros',
          'STATUS_AUDITORIA': 'Suspeita de Duplicidade'
        };
      });
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Duplicidades");
    XLSX.writeFile(wb, "Relatorio_Duplicidades_Filtradas.xlsx");
  };

  const resetFilters = () => {
    setFilters({
      prestador: '',
      beneficiario: '',
      matricula: '',
      procedimento: '',
      data: '',
      situacao: '',
      intervalCategory: ''
    });
  };

  const resetToHome = () => {
    setData(null);
    setDuplicateGroups([]);
    setFileName(null);
    setTransactionCount(0);
    resetFilters();
  };

  return (
    <div className="min-h-screen bg-[#0A0B0D] text-[#E0E0E0] font-sans flex flex-col overflow-x-hidden">
      {/* Header Section */}
      <header className="flex flex-col md:flex-row items-center justify-between px-8 py-6 border-b border-white/10 gap-6">
        <div 
          className="cursor-pointer group"
          onClick={resetToHome}
        >
          <h1 className="text-2xl font-light tracking-widest text-white uppercase group-hover:text-emerald-400 transition-colors">Auditoria de Duplicidade</h1>
          <p className="text-xs text-gray-500 mt-1 uppercase tracking-tighter italic">Sistema de Verificação de Integridade de Relatórios{fileName ? ` • ${fileName}` : ''}</p>
        </div>
        <div className="flex items-center gap-6">
          {data && (
            <button 
              onClick={resetToHome}
              className="flex items-center gap-2 px-4 py-2 text-[10px] uppercase tracking-widest text-[#E0E0E0]/60 hover:text-white border border-white/10 rounded hover:bg-white/5 transition-all"
            >
              <Trash2 size={14} /> Novo Relatório
            </button>
          )}
          <div className="flex gap-4">
            <div className="px-4 py-2 bg-white/5 border border-white/10 rounded flex flex-col min-w-[140px]">
              <span className="text-[10px] text-gray-500 uppercase">Registros Analisados</span>
              <span className="text-xl font-mono text-white leading-none">{transactionCount}</span>
            </div>
            <div className="px-4 py-2 bg-red-900/20 border border-red-500/30 rounded flex flex-col min-w-[140px]">
              <span className="text-[10px] text-red-400 uppercase">Possíveis Duplicidades</span>
              <span className="text-xl font-mono text-red-500 leading-none">{duplicateGroups.length}</span>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 p-8 flex flex-col gap-6">
        {!data && !isProcessing && (
          <motion.div 
            initial={{ opacity: 0, scale: 0.98 }}
            animate={{ opacity: 1, scale: 1 }}
            className="flex-1 flex flex-col items-center justify-center border border-white/5 rounded-2xl bg-white/[0.01]"
          >
            <div 
              onDragOver={(e) => e.preventDefault()}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
              className="group cursor-pointer flex flex-col items-center p-16 border-2 border-dashed border-white/10 hover:border-white/30 rounded-xl transition-all"
            >
              <div className="w-20 h-20 bg-white/5 rounded-full flex items-center justify-center mb-6 group-hover:scale-110 transition-transform">
                <Upload size={32} className="text-white/60 group-hover:text-white" />
              </div>
              <h2 className="text-xl font-light text-white mb-2 uppercase tracking-wide">Importar Relatório</h2>
              <p className="text-xs text-gray-500 font-mono mb-8">FORMATOS: .XLSX, .XLS, .CSV</p>
              <input 
                type="file" 
                ref={fileInputRef}
                className="hidden" 
                accept=".xlsx, .xls, .csv"
                onChange={(e) => e.target.files?.[0] && processFile(e.target.files[0])}
              />
              <button className="px-8 py-3 bg-white/10 border border-white/10 text-xs uppercase tracking-widest text-white hover:bg-white/20 transition-all">
                Explorar Arquivos
              </button>
            </div>

            <div className="mt-16 w-full max-w-4xl grid grid-cols-1 md:grid-cols-2 gap-8 text-left border-t border-white/5 pt-12">
              <div className="space-y-4">
                <h3 className="text-xs font-bold uppercase tracking-widest text-emerald-500 flex items-center gap-2">
                  <div className="w-1 h-1 bg-emerald-500 rounded-full" /> Formato de Arquivo Sugerido
                </h3>
                <p className="text-sm text-gray-400 font-light leading-relaxed">
                  Para máxima precisão, recomendamos um arquivo <span className="text-white font-medium">.XLSX</span> ou <span className="text-white font-medium">.CSV</span> com as seguintes colunas:
                </p>
                <div className="grid grid-cols-2 gap-2">
                  {[
                    'DATA', 'HORA', 'NOME_BENEFICIARIO', 'MATRICULA_BENEFICIARIO', 
                    'NOME_PRESTADOR', 'MATRICULA_PRESTADOR', 'ESPECIALIDADE',
                    'PROCEDIMENTO', 'QUANTIDADE', 'SITUACAO', 'CONSELHO', 'TIPO'
                  ].map(col => (
                    <div key={col} className="flex items-center gap-2 text-[10px] font-mono text-gray-500 bg-white/5 px-2 py-1 rounded">
                      <CheckCircle2 size={10} className="text-emerald-500/50" /> {col}
                    </div>
                  ))}
                </div>
              </div>
              <div className="space-y-4">
                <h3 className="text-xs font-bold uppercase tracking-widest text-amber-500 flex items-center gap-2">
                  <div className="w-1 h-1 bg-amber-500 rounded-full" /> Dicas de Processamento
                </h3>
                <ul className="space-y-3 text-sm text-gray-400 font-light">
                  <li className="flex items-start gap-2">
                    <span className="text-amber-500/50 mt-1">●</span>
                    <span>Diferenciamos automaticamente Matrículas de Beneficiários e Prestadores.</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-amber-500/50 mt-1">●</span>
                    <span>O sistema possui algoritmos de aliasing para tentar identificar colunas com nomes variados, mas o padrão acima é o ideal.</span>
                  </li>
                  <li className="flex items-start gap-2">
                    <span className="text-amber-500/50 mt-1">●</span>
                    <span>Campos de Matrícula e DV podem estar separados ou unificados.</span>
                  </li>
                </ul>
              </div>
            </div>

            <div className="mt-12 text-center text-[10px] text-gray-600 uppercase tracking-widest flex gap-8">
              <span>● Processamento Local</span>
              <span>● Analisador Heurístico</span>
              <span>● Exportação Direta</span>
            </div>
          </motion.div>
        )}

        {isProcessing && (
          <div className="flex-1 flex flex-col items-center justify-center">
            <div className="relative">
              <div className="w-12 h-12 border-2 border-white/10 border-t-white rounded-full animate-spin"></div>
              <Clock className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 text-white/20" size={16} />
            </div>
            <p className="mt-4 text-[10px] uppercase font-mono tracking-[0.3em] text-gray-500 animate-pulse">Sincronizando Dados</p>
          </div>
        )}

        {error && (
          <div className="p-4 bg-red-900/10 border border-red-500/20 rounded text-red-400 text-xs font-mono mb-6 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <AlertCircle size={14} />
              <span>{error}</span>
            </div>
            <button onClick={() => setError(null)} className="hover:text-red-300">
              <X size={14} />
            </button>
          </div>
        )}

        {data && !isProcessing && (
          <div className="flex flex-col gap-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
            {/* Visual Quantification & Tab Switcher Bar */}
            <div className="flex flex-col md:flex-row justify-between items-stretch md:items-center gap-4 bg-white/[0.02] border border-white/10 p-4 rounded-lg">
              <div className="flex items-center gap-3 overflow-x-auto">
                <div className="px-3 py-2 bg-white/5 border border-white/10 rounded flex items-center gap-2">
                  <span className="text-[10px] text-gray-400 uppercase">Total:</span>
                  <span className="text-sm font-mono font-bold text-white">{categoryStats.total}</span>
                </div>
                <div className="px-3 py-2 bg-red-950/40 border border-red-500/30 rounded flex items-center gap-2">
                  <span className="w-2 h-2 rounded-full bg-red-500 animate-pulse"></span>
                  <span className="text-[10px] text-red-300 uppercase">≤ 30m (Provável):</span>
                  <span className="text-sm font-mono font-bold text-red-400">{categoryStats.high}</span>
                </div>
                <div className="px-3 py-2 bg-amber-950/30 border border-amber-500/30 rounded flex items-center gap-2">
                  <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                  <span className="text-[10px] text-amber-300 uppercase">30m-1h (Menos):</span>
                  <span className="text-sm font-mono font-bold text-amber-400">{categoryStats.medium}</span>
                </div>
                <div className="px-3 py-2 bg-white/5 border border-white/10 rounded flex items-center gap-2">
                  <span className="w-2 h-2 rounded-full bg-gray-400"></span>
                  <span className="text-[10px] text-gray-400 uppercase">&gt; 1h:</span>
                  <span className="text-sm font-mono font-bold text-gray-300">{categoryStats.low}</span>
                </div>
              </div>

              <div className="flex items-center gap-2 self-end md:self-auto">
                <button
                  onClick={() => setActiveTab('table')}
                  className={cn(
                    "px-4 py-2 text-[10px] uppercase font-mono tracking-widest rounded border transition-all flex items-center gap-2",
                    activeTab === 'table' ? "bg-white/15 border-white/30 text-white" : "bg-white/5 border-white/10 text-gray-400 hover:text-white"
                  )}
                >
                  <Search size={14} /> Listagem ({filteredDuplicates.length})
                </button>
                <button
                  onClick={() => setActiveTab('dashboard')}
                  className={cn(
                    "px-4 py-2 text-[10px] uppercase font-mono tracking-widest rounded border transition-all flex items-center gap-2",
                    activeTab === 'dashboard' ? "bg-white/15 border-white/30 text-emerald-400 border-emerald-500/40" : "bg-white/5 border-white/10 text-gray-400 hover:text-white"
                  )}
                >
                  <BarChart3 size={14} /> Dashboard & Picos
                </button>
              </div>
            </div>

            {activeTab === 'dashboard' ? (
              <div className="flex flex-col gap-6 animate-in fade-in duration-300">
                {/* Dashboard Header & Export */}
                <div className="flex justify-between items-center bg-white/[0.02] border border-white/10 px-6 py-4 rounded-lg">
                  <div>
                    <h3 className="text-sm font-bold uppercase tracking-widest text-white flex items-center gap-2">
                      <BarChart3 size={16} className="text-emerald-400" /> Dashboard Executivo de Duplicidades e Picos
                    </h3>
                    <p className="text-xs text-gray-400 mt-0.5">Análise detalhada de horários, dias de pico, prestadores e indicadores críticos de recorrência.</p>
                  </div>
                  <button 
                    onClick={exportDashboardReport}
                    className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-[10px] uppercase font-bold tracking-widest rounded transition-all flex items-center gap-2 shadow-lg shadow-emerald-900/30"
                  >
                    <FileDown size={14} /> Exportar Relatório do Dashboard (.XLSX)
                  </button>
                </div>

                {/* Peak Analysis Grid */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {/* Peak Days */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <h4 className="text-xs font-bold uppercase tracking-widest text-gray-300 flex items-center gap-2">
                      <Calendar size={14} className="text-emerald-400" /> Dias com Maior Incidência de Duplicidades
                    </h4>
                    <div className="space-y-3 max-h-[300px] overflow-y-auto custom-scrollbar pr-2">
                      {dashboardStats.topDays.length === 0 ? (
                        <p className="text-xs text-gray-500 italic">Nenhum dado de data disponível.</p>
                      ) : (
                        dashboardStats.topDays.map((item) => {
                          const maxCount = dashboardStats.topDays[0]?.count || 1;
                          const pct = Math.round((item.count / maxCount) * 100);
                          return (
                            <div key={item.day} className="flex flex-col gap-1 bg-white/5 p-3 rounded border border-white/5">
                              <div className="flex justify-between items-center text-xs">
                                <span className="font-mono text-white font-medium">{item.day || 'Data Indefinida'}</span>
                                <span className="font-mono font-bold text-emerald-400">{item.count} grupo(s)</span>
                              </div>
                              <div className="w-full bg-white/10 h-1.5 rounded-full overflow-hidden">
                                <div className="bg-emerald-500 h-full rounded-full transition-all duration-500" style={{ width: `${pct}%` }} />
                              </div>
                            </div>
                          );
                        })
                      )}
                    </div>
                  </div>

                  {/* Peak Hours */}
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <h4 className="text-xs font-bold uppercase tracking-widest text-gray-300 flex items-center gap-2">
                      <Clock size={14} className="text-amber-400" /> Horários de Pico (Faixas Horárias)
                    </h4>
                    <div className="space-y-3 max-h-[300px] overflow-y-auto custom-scrollbar pr-2">
                      {dashboardStats.topHours.length === 0 ? (
                        <p className="text-xs text-gray-500 italic">Nenhum horário registrado.</p>
                      ) : (
                        dashboardStats.topHours.map((item) => {
                          const maxCount = dashboardStats.topHours[0]?.count || 1;
                          const pct = Math.round((item.count / maxCount) * 100);
                          return (
                            <div key={item.hour} className="flex flex-col gap-1 bg-white/5 p-3 rounded border border-white/5">
                              <div className="flex justify-between items-center text-xs">
                                <span className="font-mono text-white font-medium">{item.hour}</span>
                                <span className="font-mono font-bold text-amber-400">{item.count} ocorrência(s)</span>
                              </div>
                              <div className="w-full bg-white/10 h-1.5 rounded-full overflow-hidden">
                                <div className="bg-amber-500 h-full rounded-full transition-all duration-500" style={{ width: `${pct}%` }} />
                              </div>
                            </div>
                          );
                        })
                      )}
                    </div>
                  </div>
                </div>

                {/* Top Providers & Procedures */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <h4 className="text-xs font-bold uppercase tracking-widest text-gray-300 flex items-center gap-2">
                      <Building2 size={14} className="text-blue-400" /> Prestadores com Mais Duplicidades
                    </h4>
                    <div className="space-y-2">
                      {dashboardStats.topProviders.map((p, idx) => (
                        <div key={idx} className="flex justify-between items-center text-xs bg-white/5 p-3 rounded border border-white/5">
                          <span className="text-white truncate max-w-[280px]" title={p.provider}>{p.provider}</span>
                          <span className="font-mono font-bold text-blue-400">{p.count}</span>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="bg-white/[0.02] border border-white/10 p-6 rounded-lg flex flex-col gap-4">
                    <h4 className="text-xs font-bold uppercase tracking-widest text-gray-300 flex items-center gap-2">
                      <FileText size={14} className="text-purple-400" /> Procedimentos Mais Frequentes em Duplicidade
                    </h4>
                    <div className="space-y-2">
                      {dashboardStats.topProcedures.map((pr, idx) => (
                        <div key={idx} className="flex justify-between items-center text-xs bg-white/5 p-3 rounded border border-white/5">
                          <span className="text-white truncate max-w-[280px]" title={pr.procedure}>{pr.procedure}</span>
                          <span className="font-mono font-bold text-purple-400">{pr.count}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            ) : (
              <>
                {/* Filter Section */}
                <section className="bg-white/5 border border-white/10 p-6 rounded-lg flex flex-col gap-6">
                  <div className="flex items-center justify-between border-b border-white/5 pb-4">
                    <div className="flex items-center gap-3">
                      <Filter size={16} className="text-gray-500" />
                      <h3 className="text-xs uppercase font-bold tracking-widest text-gray-400">Filtros de Auditoria</h3>
                    </div>
                    {Object.values(filters).some(v => v !== '') && (
                      <button 
                        onClick={resetFilters}
                        className="text-[10px] uppercase text-gray-500 hover:text-white flex items-center gap-1 transition-colors"
                      >
                        <Trash2 size={12} /> Limpar Filtros
                      </button>
                    )}
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-8 gap-4">
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Beneficiário</label>
                      <div className="relative">
                        <User className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-600" size={14} />
                        <input 
                          value={filters.beneficiario}
                          onChange={e => setFilters(prev => ({ ...prev, beneficiario: e.target.value }))}
                          className="w-full bg-[#0A0B0D] border border-white/10 rounded px-9 py-2 text-xs focus:border-white/30 outline-none transition-all"
                          placeholder="Nome..."
                        />
                      </div>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Matrícula (Bene/Prest)</label>
                      <div className="relative">
                        <CheckCircle2 className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-600" size={14} />
                        <input 
                          value={filters.matricula}
                          onChange={e => setFilters(prev => ({ ...prev, matricula: e.target.value }))}
                          className="w-full bg-[#0A0B0D] border border-white/10 rounded px-9 py-2 text-xs focus:border-white/30 outline-none transition-all"
                          placeholder="Nº..."
                        />
                      </div>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Prestador</label>
                      <select 
                        value={filters.prestador}
                        onChange={e => setFilters(prev => ({ ...prev, prestador: e.target.value }))}
                        className="w-full bg-[#0A0B0D] border border-white/10 rounded px-4 py-2 text-xs focus:border-white/30 outline-none transition-all cursor-pointer"
                      >
                        <option value="">Todos</option>
                        {uniqueValues.prestadores.map(p => (
                          <option key={p} value={p}>{p}</option>
                        ))}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">CRM Executante</label>
                      <select 
                        value={filters.conselhoExecutante}
                        onChange={e => setFilters(prev => ({ ...prev, conselhoExecutante: e.target.value }))}
                        className="w-full bg-[#0A0B0D] border border-white/10 rounded px-4 py-2 text-xs focus:border-white/30 outline-none transition-all cursor-pointer"
                      >
                        <option value="">Todos</option>
                        {uniqueValues.conselhoExecutantes.map(c => (
                          <option key={c} value={c}>{c}</option>
                        ))}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Procedimento</label>
                      <select 
                        value={filters.procedimento}
                        onChange={e => setFilters(prev => ({ ...prev, procedimento: e.target.value }))}
                        className="w-full bg-[#0A0B0D] border border-white/10 rounded px-4 py-2 text-xs focus:border-white/30 outline-none transition-all cursor-pointer"
                      >
                        <option value="">Todos</option>
                        {uniqueValues.procedimentos.map(p => (
                          <option key={p} value={p}>{p}</option>
                        ))}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Data</label>
                      <select 
                        value={filters.data}
                        onChange={e => setFilters(prev => ({ ...prev, data: e.target.value }))}
                        className="w-full bg-[#0A0B0D] border border-white/10 rounded px-4 py-2 text-xs focus:border-white/30 outline-none transition-all cursor-pointer"
                      >
                        <option value="">Todas</option>
                        {uniqueValues.datas.map(d => (
                          <option key={d} value={d}>{d}</option>
                        ))}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Situação</label>
                      <select 
                        value={filters.situacao}
                        onChange={e => setFilters(prev => ({ ...prev, situacao: e.target.value }))}
                        className="w-full bg-[#0A0B0D] border border-white/10 rounded px-4 py-2 text-xs focus:border-white/30 outline-none transition-all cursor-pointer"
                      >
                        <option value="">Todas</option>
                        {uniqueValues.situacoes.map(s => (
                          <option key={s} value={s}>{s}</option>
                        ))}
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] uppercase text-gray-500 font-mono">Intervalo (Probabilidade)</label>
                      <select 
                        value={filters.intervalCategory}
                        onChange={e => setFilters(prev => ({ ...prev, intervalCategory: e.target.value }))}
                        className="w-full bg-[#0A0B0D] border border-white/10 rounded px-4 py-2 text-xs focus:border-white/30 outline-none transition-all cursor-pointer"
                      >
                        <option value="">Todos</option>
                        <option value="high">≤ 30min (Bem Provável)</option>
                        <option value="medium">30min - 1h (Pouco Menos)</option>
                        <option value="low">&gt; 1h</option>
                      </select>
                    </div>
                  </div>
                </section>

            {/* Actions Bar */}
            <div className="flex justify-between items-center bg-white/[0.02] border border-white/10 px-6 py-3 rounded-lg">
              <span className="text-[10px] text-gray-500 uppercase tracking-widest font-mono">
                {filteredDuplicates.length} grupos de duplicidades encontrados com os filtros atuais
              </span>
              <div className="flex gap-4">
                <button 
                   disabled={filteredDuplicates.length === 0}
                   onClick={exportResults}
                   className="px-4 py-1.5 bg-white/10 border border-white/20 rounded text-[10px] uppercase tracking-widest text-white hover:bg-white/20 transition-all flex items-center gap-2 disabled:opacity-20 disabled:cursor-not-allowed"
                >
                  <FileDown size={14} /> Exportar XLSX
                </button>
              </div>
            </div>

            {/* Results Table */}
            <div className="bg-white/[0.02] border border-white/10 rounded-lg overflow-hidden flex flex-col shadow-2xl">
              <div className="grid grid-cols-12 bg-white/5 py-3 px-4 text-[10px] uppercase font-bold tracking-widest text-gray-400 border-b border-white/10">
                <div className="col-span-2">Beneficiário</div>
                <div className="col-span-2">Prestador</div>
                <div className="col-span-1">Especialidade</div>
                <div className="col-span-2">Procedimento</div>
                <div className="col-span-1 text-center">Tipo</div>
                <div className="col-span-1 text-center">Data</div>
                <div className="col-span-1 text-center">Qtd</div>
                <div className="col-span-1 text-center">Situação</div>
                <div className="col-span-1 text-center text-red-400 bg-red-400/5 rounded">Horários</div>
              </div>

              <div className="divide-y divide-white/5 max-h-[600px] overflow-y-auto custom-scrollbar">
                {filteredDuplicates.length === 0 ? (
                  <div className="py-24 flex flex-col items-center justify-center opacity-40">
                    <Search size={32} className="mb-4" />
                    <p className="text-sm font-light uppercase tracking-[0.2em]">Nenhum registro corresponde aos filtros</p>
                  </div>
                ) : (
                  filteredDuplicates.map((group, idx) => (
                    <div 
                      key={group.fingerprint}
                      className="grid grid-cols-12 items-center py-4 px-4 hover:bg-white/[0.03] transition-colors"
                    >
                      <div className="col-span-2 pr-4 border-r border-white/5">
                        <div className="flex flex-col gap-1 overflow-hidden">
                          <p className="text-xs text-white font-medium truncate uppercase" title={group.params.beneficiario}>{group.params.beneficiario}</p>
                          <span className="text-[8px] text-emerald-500/60 font-mono self-start bg-emerald-500/5 px-1 rounded border border-emerald-500/10" title="Matrícula Beneficiário">
                            {group.params.beneficiario_matricula || '---'}
                          </span>
                        </div>
                      </div>
                      <div className="col-span-2 px-4 border-r border-white/5">
                        <div className="flex flex-col gap-1 overflow-hidden">
                          <p className="text-xs text-gray-300 truncate uppercase" title={group.params.prestador}>{group.params.prestador}</p>
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className="text-[8px] text-blue-500/60 font-mono self-start bg-blue-500/5 px-1 rounded border border-blue-500/10" title="Matrícula Prestador">
                              {group.params.prestador_matricula || '---'}
                            </span>
                            {group.params.conselhoExecNum && (
                              <span className="text-[8px] text-purple-400/80 font-mono self-start bg-purple-500/5 px-1 rounded border border-purple-500/10" title="CRM Executante">
                                {group.params.conselhoExecTipo || 'CRM'} {group.params.conselhoExecNum}{group.params.conselhoExecUf ? `/${group.params.conselhoExecUf}` : ''}
                              </span>
                            )}
                          </div>
                        </div>
                      </div>
                      <div className="col-span-1 text-[9px] text-gray-400 text-center truncate italic px-1" title={group.params.especialidade}>
                        {group.params.especialidade || ''}
                      </div>
                      <div className="col-span-2 text-xs truncate font-mono text-white/90 px-2" title={group.params.procedimento}>
                        {group.params.procedimento}
                      </div>
                      <div className="col-span-1 text-[9px] text-blue-400 text-center truncate uppercase font-mono px-1" title={group.params.tipo}>
                        {group.params.tipo || ''}
                      </div>
                      <div className="col-span-1 text-xs text-gray-400 text-center">
                        {group.params.data}
                      </div>
                      <div className="col-span-1 text-xs font-mono text-white/80 text-center font-bold">
                        {group.params.quantidade.padStart(2, '0')}
                      </div>
                      <div className="col-span-1">
                        <span className={cn(
                          "text-[9px] uppercase px-2 py-0.5 rounded-full inline-block font-bold",
                          group.params.situacao.toLowerCase() === 'aprovada' 
                            ? "bg-emerald-500/10 text-emerald-500" 
                            : "bg-white/10 text-gray-400"
                        )}>
                          {group.params.situacao}
                        </span>
                      </div>
                      <div className="col-span-1 flex flex-col items-center gap-1">
                        {(() => {
                          const intervalInfo = getGroupTimeInterval(group.rows);
                          return (
                            <>
                              {group.rows.map((row, rIdx) => {
                                const dateVal = getColValue(row, 'DATA');
                                const hourVal = getColValue(row, 'HORA');
                                
                                let { time } = normalizeDate(dateVal);
                                
                                // Se não encontrou hora na data, tenta na coluna HORA
                                if ((!time || time === '00:00') && hourVal !== undefined && hourVal !== null && hourVal !== '') {
                                  const hNorm = normalizeDate(hourVal);
                                  if (hNorm.time) {
                                    time = hNorm.time;
                                  } else if (typeof hourVal === 'string' && hourVal.includes(':')) {
                                    time = hourVal.trim().substring(0, 5);
                                  }
                                }

                                const badgeStyle = 
                                  intervalInfo.level === 'high' 
                                    ? "bg-red-500/25 border-red-500/50 text-red-300 font-bold shadow-sm shadow-red-500/20"
                                    : intervalInfo.level === 'medium'
                                    ? "bg-amber-500/20 border-amber-500/40 text-amber-300 font-semibold"
                                    : "bg-white/10 border-white/20 text-gray-300";
                                
                                return (
                                  <span 
                                    key={rIdx}
                                    className={cn("px-1.5 py-0.5 border rounded font-mono text-[10px] w-full text-center", badgeStyle)}
                                    title={
                                      intervalInfo.level === 'high' 
                                        ? 'Intervalo ≤ 30 min (Bem Provável)' 
                                        : intervalInfo.level === 'medium' 
                                        ? 'Intervalo entre 30 min e 1h (Um pouco menos provável)' 
                                        : 'Intervalo > 1h'
                                    }
                                  >
                                    {(time && time !== '0') ? time : '--:--'}
                                  </span>
                                );
                              })}
                              {intervalInfo.minDiff !== null && (
                                <span className={cn(
                                  "text-[8px] font-mono mt-0.5 px-1 rounded uppercase tracking-wider text-center w-full truncate",
                                  intervalInfo.level === 'high' ? "text-red-400 bg-red-400/10 font-bold" :
                                  intervalInfo.level === 'medium' ? "text-amber-400 bg-amber-400/10" : "text-gray-400 bg-white/5"
                                )} title={intervalInfo.label}>
                                  {intervalInfo.level === 'high' ? '≤30m (Provável)' : intervalInfo.level === 'medium' ? '30m-1h' : '>1h'}
                                </span>
                              )}
                            </>
                          );
                        })()}
                      </div>
                    </div>
                  ))
                )}
              </div>

              {/* Info Bar at Bottom */}
              <div className="p-4 bg-white/5 text-[10px] text-gray-400 border-t border-white/10 flex flex-col md:flex-row justify-between items-center gap-4">
                <div className="flex items-center gap-3 flex-wrap">
                  <span className="text-amber-500">●</span>
                  <span>Critério: Mesma Data, Prestador, Especialidade, Beneficiário, Procedimento, Tipo, Situação, Qtd e Conselho.</span>
                  <div className="flex items-center gap-3 ml-2 pl-3 border-l border-white/10">
                    <span className="flex items-center gap-1 text-red-400 font-medium"><span className="w-2 h-2 rounded-full bg-red-500"></span> ≤ 30min (Bem Provável)</span>
                    <span className="flex items-center gap-1 text-amber-400 font-medium"><span className="w-2 h-2 rounded-full bg-amber-500"></span> 30min-1h (Pouco Menos)</span>
                    <span className="flex items-center gap-1 text-gray-400 font-medium"><span className="w-2 h-2 rounded-full bg-gray-500"></span> &gt; 1h</span>
                  </div>
                </div>
                <div className="italic opacity-60">
                  Total de {filteredDuplicates.length} incidentes identificados para conferência.
                </div>
              </div>
            </div>
          </>
        )}
      </div>
    )}
  </main>

      {/* Internal Scrollbar Styling */}
      <style dangerouslySetInnerHTML={{ __html: `
        .custom-scrollbar::-webkit-scrollbar {
          width: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: rgba(255, 255, 255, 0.02);
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: rgba(255, 255, 255, 0.1);
          border-radius: 3px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: rgba(255, 255, 255, 0.2);
        }
      `}} />

      {/* Footer Details */}
      <footer className="px-8 py-6 border-t border-white/10 flex flex-col md:flex-row justify-between items-center bg-[#07080A] gap-4">
        <div className="flex items-center gap-2">
          <div className={cn("w-2 h-2 rounded-full", data ? "bg-emerald-500" : "bg-gray-700")}></div>
          <span className="text-[10px] text-gray-500 uppercase tracking-widest">
            AuditCheck System — Status: {data ? "Processamento Concluído" : "Pronto para Importação"}
          </span>
        </div>
        <div className="text-[10px] text-gray-500 font-mono uppercase tracking-tighter">
          Integridade e Rigor na Auditoria de Transações HealthTech
        </div>
      </footer>
    </div>
  );
}

