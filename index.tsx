import React from 'react';
import { createRoot } from 'react-dom/client';


declare var XLSX: any;
declare var Chart: any;
declare var JSZip: any;
declare var marked: any;

const { useState, useEffect, useRef, useCallback, useMemo } = React;

// FIX: Add type definitions for data structures
interface SalesRecord {
    'Date'?: string | Date;
    'Transaction ID'?: string | number;
    'Product'?: string;
    'Category'?: string;
    'Region'?: string;
    'Quantity'?: number;
    'Unit Price'?: number;
    'Total Revenue'?: number;
    [key: string]: any;
}

interface CleanedSalesData {
    'Date': Date;
    'Transaction ID'?: string | number;
    'Product': string;
    'Category'?: string;
    'Region'?: string;
    'Quantity'?: number;
    'Unit Price'?: number;
    'Total Revenue'?: number;
}

interface DashboardData {
    totalRevenue: number;
    averageTicket: number;
    bestMonth: string;
    totalTransactions: number;
    monthlySalesChart: { labels: string[]; data: number[]; };
    categorySalesChart: { labels: string[]; data: number[]; };
    recentSales: CleanedSalesData[];
}

interface RawDataRow {
    [key: string]: any;
}

// FIX: Add type for grounding chunk to avoid property access on unknown.
interface GroundingChunk {
    web?: {
        uri?: string;
        title?: string;
    };
    [key: string]: any; // Allow other properties like 'maps'
}


// SVG Icon Components
const IconLogo = () => (
    <svg width="28" height="28" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
);
const IconDashboard = () => <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="7" height="7"></rect><rect x="14" y="3" width="7" height="7"></rect><rect x="14" y="14" width="7" height="7"></rect><rect x="3" y="14" width="7" height="7"></rect></svg>;
const IconSun = () => <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="5"></circle><line x1="12" y1="1" x2="12" y2="3"></line><line x1="12" y1="21" x2="12" y2="23"></line><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line><line x1="1" y1="12" x2="3" y2="12"></line><line x1="21" y1="12" x2="23" y2="12"></line><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line></svg>;
const IconMoon = () => <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path></svg>;
const IconMenu = () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="3" y1="12" x2="21" y2="12"></line><line x1="3" y1="6" x2="21" y2="6"></line><line x1="3" y1="18" x2="21" y2="18"></line></svg>;
const IconUpload = () => <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="17 8 12 3 7 8"></polyline><line x1="12" y1="3" x2="12" y2="15"></line></svg>;
const IconChat = () => <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="currentColor"><path d="M20 2H4c-1.1 0-2 .9-2 2v18l4-4h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2z"></path></svg>;


const App = () => {
    const [theme, setTheme] = useState('dark');
    const [isSidebarOpen, setSidebarOpen] = useState(false);
    const [allData, setAllData] = useState<CleanedSalesData[]>([]);
    const [filteredData, setFilteredData] = useState<DashboardData | null>(null);
    const [currentFilteredRawData, setCurrentFilteredRawData] = useState<CleanedSalesData[]>([]);
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState('');
    const fileInputRef = useRef<HTMLInputElement>(null);
    const [filters, setFilters] = useState({ month: 'all', category: 'all', region: 'all' });
    const [filterOptions, setFilterOptions] = useState<{months: string[], categories: string[], regions: string[]}>({ months: [], categories: [], regions: [] });
    
    const [isChatOpen, setChatOpen] = useState(false);
    const [chatMessages, setChatMessages] = useState<{ sender: 'user' | 'bot'; text: string; }[]>([]);
    const [isChatLoading, setChatLoading] = useState(false);
    const [chat, setChat] = useState<Chat | null>(null);

    const suggestedQuestions = [
        "Qual foi o produto mais vendido?",
        "Compare as vendas entre as regiões.",
        "Qual a tendência de receita?",
        "Qual categoria teve o menor desempenho?"
    ];

    useEffect(() => {
        document.body.setAttribute('data-theme', theme);
    }, [theme]);

    const normalizeJsonData = (jsonData: RawDataRow[]): SalesRecord[] => {
        return jsonData.map(row => {
            const normalizedRow: SalesRecord = {};
            const keyMap: { [key: string]: keyof SalesRecord } = {
                'date': 'Date', 'data': 'Date',
                'id_transacao': 'Transaction ID', 'id transacao': 'Transaction ID',
                'product': 'Product', 'produto': 'Product',
                'category': 'Category', 'categoria': 'Category',
                'region': 'Region', 'região': 'Region', 'regiao': 'Region',
                'quantity': 'Quantity', 'quantidade': 'Quantity', 'qtd': 'Quantity',
                'unit price': 'Unit Price', 'preço unitário': 'Unit Price', 'preco unitario': 'Unit Price', 'unitprice': 'Unit Price', 'preço_unitário': 'Unit Price',
                'total revenue': 'Total Revenue', 'receita total': 'Total Revenue', 'receita_total': 'Total Revenue'
            };
            for (const key in row) {
                if (Object.prototype.hasOwnProperty.call(row, key)) {
                    const normalizedKey = key.trim().toLowerCase();
                    if (keyMap[normalizedKey]) {
                        normalizedRow[keyMap[normalizedKey]] = row[key];
                    }
                }
            }
            return normalizedRow;
        });
    };
    
    const handleFile = (file: File, handler: (data: ArrayBuffer, fileName: string) => SalesRecord[] | Promise<SalesRecord[]>) => new Promise<SalesRecord[]>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => {
            try {
                const result = handler(e.target!.result as ArrayBuffer, file.name);
                resolve(result);
            } catch (err) {
                reject(err);
            }
        };
        // FIX: Add type to file parameter, fixing error on file.name
        reader.onerror = err => reject(new Error(`Erro ao ler o arquivo ${file.name}.`));
        reader.readAsArrayBuffer(file);
    });

    const processSheetData = (data: ArrayBuffer, fileName: string): SalesRecord[] => {
        try {
            // FIX: Force UTF-8 codepage to correctly handle accented characters.
            const workbook = XLSX.read(new Uint8Array(data), { type: 'array', cellDates: true, codepage: 65001 });
            const jsonData: RawDataRow[] = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
            return normalizeJsonData(jsonData);
        } catch (err) {
            throw new Error(`Erro ao processar ${fileName}.`);
        }
    };

    const processZipData = async (data: ArrayBuffer, fileName: string): Promise<SalesRecord[]> => {
        try {
            const zip = await JSZip.loadAsync(data);
            const filePromises: Promise<RawDataRow[]>[] = [];
            zip.forEach((path, entry) => {
                if (!entry.dir && (path.toLowerCase().endsWith('.csv') || path.toLowerCase().endsWith('.xlsx'))) {
                    filePromises.push(entry.async('array').then(content => {
                        // FIX: Force UTF-8 codepage to correctly handle accented characters within ZIP files.
                        const workbook = XLSX.read(content, { type: 'array', cellDates: true, codepage: 65001 });
                        return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
                    }));
                }
            });
            return normalizeJsonData((await Promise.all(filePromises)).flat());
        } catch (err) {
            throw new Error(`Erro ao processar o ZIP ${fileName}.`);
        }
    };
    
    const handleFileUpload = useCallback(async (files: FileList | null) => {
        if (!files || files.length === 0) return;
        setIsLoading(true);
        setError('');
        setAllData([]);
        setFilteredData(null);
        setFilters({ month: 'all', category: 'all', region: 'all' });
        
        try {
            const promises: Promise<SalesRecord[]>[] = Array.from(files).map(file => {
                const ext = file.name.toLowerCase().split('.').pop() || '';
                if (ext === 'zip') return handleFile(file, processZipData);
                if (['xlsx', 'csv'].includes(ext)) return handleFile(file, processSheetData);
                return Promise.resolve([] as SalesRecord[]);
            });
            const data: SalesRecord[] = (await Promise.all(promises)).flat();
            if (data.length === 0) throw new Error("Nenhum dado válido encontrado.");

            // FIX: Add types to 'r' to fix property access and spread operator errors
            const cleanData: CleanedSalesData[] = data
                .filter((r): r is SalesRecord & { Date: string | Date | number, Product: string } => !!r.Date && !!r.Product)
                .map(r => ({ ...r, Date: new Date(r.Date)}));
            setAllData(cleanData);
        } catch (err: any) {
            setError(err.message);
        } finally {
            setIsLoading(false);
        }
    }, []);

    const processDashboardData = (data: CleanedSalesData[]): DashboardData | null => {
        if (!data || data.length === 0) return null;
        
        let totalRevenue = 0;
        const monthlySales: { [key: string]: number } = {};
        const categorySales: { [key: string]: number } = {};
        let validRowCount = 0;

        data.forEach(row => {
            const { Date: date, Category, Quantity, 'Unit Price': unitPrice, 'Total Revenue': totalRevenueFromRow } = row;
            if (isNaN(date.getTime())) return;
            
            let saleAmount = Number(totalRevenueFromRow);
            if (isNaN(saleAmount) || saleAmount === 0) {
                 saleAmount = Number(Quantity) * Number(unitPrice);
            }
            if (isNaN(saleAmount)) return;

            validRowCount++;
            totalRevenue += saleAmount;
            
            const monthYearKey = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
            monthlySales[monthYearKey] = (monthlySales[monthYearKey] || 0) + saleAmount;
            
            if (Category) {
                if (!categorySales[Category]) categorySales[Category] = 0;
                categorySales[Category] += saleAmount;
            }
        });

        if (validRowCount === 0) return null;

        const sortedMonths = Object.entries(monthlySales).sort((a, b) => a[0].localeCompare(b[0]));
        const bestMonthEntry = Object.entries(monthlySales).reduce((a, b) => a[1] > b[1] ? a : b, ['', -Infinity]);
        
        const formatMonthYear = (key: string) => {
            if (!key) return 'N/A';
            const [year, month] = key.split('-');
            return new Date(Number(year), Number(month) - 1).toLocaleDateString('pt-BR', { year: 'numeric', month: 'short' });
        };
        
        return {
            totalRevenue,
            averageTicket: totalRevenue / validRowCount,
            bestMonth: formatMonthYear(bestMonthEntry[0]),
            totalTransactions: validRowCount,
            monthlySalesChart: {
                labels: sortedMonths.map(item => formatMonthYear(item[0])),
                data: sortedMonths.map(item => item[1]),
            },
            categorySalesChart: {
                labels: Object.keys(categorySales),
                data: Object.values(categorySales),
            },
            recentSales: data.sort((a, b) => b.Date.getTime() - a.Date.getTime()).slice(0, 10),
        };
    };

    useEffect(() => {
        if (allData.length > 0) {
             const months = [...new Set(allData.map(r => `${r.Date.getFullYear()}-${(r.Date.getMonth() + 1).toString().padStart(2, '0')}`))].sort();
             const categories = [...new Set(allData.map(r => r.Category).filter(Boolean) as string[])].sort();
             const regions = [...new Set(allData.map(r => r.Region).filter(Boolean) as string[])].sort();
             setFilterOptions({ months, categories, regions });

             let dataToProcess = allData;
             if (filters.month !== 'all') {
                dataToProcess = dataToProcess.filter(r => `${r.Date.getFullYear()}-${(r.Date.getMonth() + 1).toString().padStart(2, '0')}` === filters.month);
             }
             if (filters.category !== 'all') {
                dataToProcess = dataToProcess.filter(r => r.Category === filters.category);
             }
             if (filters.region !== 'all') {
                 dataToProcess = dataToProcess.filter(r => r.Region === filters.region);
             }
             
             setCurrentFilteredRawData(dataToProcess);
             setFilteredData(processDashboardData(dataToProcess));
        } else {
            setCurrentFilteredRawData([]);
            setFilteredData(null);
        }
    }, [allData, filters]);

    useEffect(() => {
        if (!filteredData) {
            setChat(null);
            return;
        }

        if (!process.env.API_KEY) {
            console.error("Chave de API do Google não configurada. As funcionalidades de IA estão desativadas.");
            setChat(null);
            return;
        }

        const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

        const currencyFormatter = (value: number) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value || 0);
        const numberFormatter = (value: number) => new Intl.NumberFormat('pt-BR').format(value || 0);

        // --- START OF CONCISE SUMMARIZATION LOGIC ---

        const productSales: { [key: string]: number } = {};
        currentFilteredRawData.forEach(row => {
            let saleAmount = Number(row['Total Revenue']);
            if (isNaN(saleAmount) || saleAmount === 0) {
                saleAmount = Number(row.Quantity) * Number(row['Unit Price']);
            }
            if (isNaN(saleAmount)) return;
            if (row.Product) {
                productSales[row.Product] = (productSales[row.Product] || 0) + saleAmount;
            }
        });
        const topProducts = Object.entries(productSales)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 5)
            .map(([name, revenue]) => `${name}: ${currencyFormatter(revenue)}`)
            .join('; ');

        const regionSales: { [key: string]: number } = {};
        currentFilteredRawData.forEach(row => {
            if (row.Region) {
                let saleAmount = Number(row['Total Revenue']);
                if (isNaN(saleAmount) || saleAmount === 0) {
                    saleAmount = Number(row.Quantity) * Number(row['Unit Price']);
                }
                if (isNaN(saleAmount)) return;
                regionSales[row.Region] = (regionSales[row.Region] || 0) + saleAmount;
            }
        });
        const topRegionsSummary = Object.entries(regionSales)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 10)
            .map(([name, revenue]) => `${name}: ${currencyFormatter(revenue)}`)
            .join('; ');
        
        const categoryData = filteredData.categorySalesChart.labels.map((label, index) => ({
            label, value: filteredData.categorySalesChart.data[index]
        }));
        const topCategoriesSummary = categoryData.sort((a, b) => b.value - a.value).slice(0, 10).map(item => `${item.label}: ${currencyFormatter(item.value)}`).join('; ');

        const monthlyData = filteredData.monthlySalesChart.labels.map((label, index) => ({
            label, value: filteredData.monthlySalesChart.data[index]
        }));
        const recentMonthsSummary = monthlyData.slice(-12).map(item => `${item.label}: ${currencyFormatter(item.value)}`).join('; ');

        const dataSummary = `
            - Receita Total: ${currencyFormatter(filteredData.totalRevenue)}
            - Ticket Médio: ${currencyFormatter(filteredData.averageTicket)}
            - Mês de Pico de Vendas: ${filteredData.bestMonth}
            - Total de Transações: ${numberFormatter(filteredData.totalTransactions)}
            - Vendas por Categoria (Top 10): ${topCategoriesSummary || 'N/A'}
            - Vendas Mensais (Últimos 12 meses): ${recentMonthsSummary || 'N/A'}
            - Vendas por Região (Top 10): ${topRegionsSummary || 'N/A'}
            - Top 5 Produtos por Receita: ${topProducts || 'N/A'}
        `;
        // --- END OF CONCISE SUMMARIZATION LOGIC ---
        
        const systemInstruction = `**Identidade e Contexto:**
Você é o **Assistente de Dados da Alpha Insights**, uma IA especializada em análises de performance e inteligência de negócios para o setor de varejo de tecnologia.

**Sua Missão:**
Atuar como um analista virtual inteligente, fornecendo insights precisos e acionáveis sobre vendas, produtos, categorias e receitas.

**Regras de Comunicação e Estilo (Cruciais):**

1.  **Tom e Linguagem:**
    * **Profissional, Clara e Empática.**
    * **EVITE JARGÕES TÉCNICOS** e respostas robóticas.
    * Prefira **frases curtas e objetivas**.

2.  **Precisão e Interpretação:**
    * Responda com clareza e precisão, interpretando corretamente os dados fornecidos no resumo.
    * **Corrija automaticamente** erros de digitação e grafias incorretas na entrada do usuário (ex: "receita" em vez de "resceita").

3.  **Explicação de Dados:**
    * Sempre que citar números (percentuais, totais, médias), **EXPLIQUE O SIGNIFICADO**.
        * *Exemplo:* "Um crescimento de 12% representa um aumento considerável nas vendas no comparativo mensal."
    * Quando apropriado, **resuma visualmente em tópicos ou bullet points**.

4.  **Fonte dos Dados e Busca Externa:**
    * **PRIORIDADE 1: Análise Interna.** Sempre tente responder à pergunta usando **PRIMEIRO** o **"RESUMO DOS DADOS ATUAIS"** fornecido abaixo.
    * **PRIORIDADE 2: Busca Externa.** Se a informação solicitada **NÃO ESTIVER** no resumo de dados, utilize a ferramenta de busca do Google para encontrar a resposta.
    * **TRANSPARÊNCIA:** Ao usar a busca externa, **informe claramente ao usuário**. Por exemplo: "Nos dados carregados, não encontrei essa informação. Com base em uma consulta externa,..."
    * **NÃO INVENTE INFORMAÇÕES.** Se não conseguir encontrar a resposta nem nos dados internos nem na busca externa, informe que não foi possível encontrar a informação.

**Exemplos de Análises que você pode fazer com os dados do resumo:**
- **Variação percentual de receita entre meses:** Com base nos dados de 'Vendas Mensais'.
- **Comparar vendas entre regiões:** Com base em 'Vendas por Região'.
- **Identificar os produtos com maior receita:** Com base nos 'Top 5 Produtos por Receita'.
- **Tendência geral de crescimento:** Analisando a progressão das 'Vendas Mensais'.

**Instrução Final:**
Implemente as regras acima e use o resumo de dados para realizar análises imediatas. Se necessário, recorra à busca externa para complementar sua resposta, mantendo sempre a transparência sobre a origem da informação.

---
**RESUMO DOS DADOS ATUAIS PARA ANÁLISE:**
${dataSummary}
---
`;
        const newChat = ai.chats.create({
            model: 'gemini-2.5-flash',
            config: {
                systemInstruction: systemInstruction,
                tools: [{googleSearch: {}}],
            },
        });
        setChat(newChat);

    }, [filteredData, currentFilteredRawData]);

    const handleFilterChange = (filterType: 'month' | 'category' | 'region', value: string) => {
        setFilters(prev => ({...prev, [filterType]: value}));
    };
    
    const handleSendMessage = async (message: string) => {
        if (!message.trim() || isChatLoading || !chat) return;

        const userMessage = { sender: 'user' as const, text: message };
        setChatMessages(prev => [...prev, userMessage]);
        setChatLoading(true);

        try {
            const response = await fetch("/.netlify/functions/chatbot", {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({ message }),
});

const data = await response.json();
const botMessage = { sender: 'bot' as const, text: data.reply || "Não foi possível gerar uma resposta." };
setChatMessages(prev => [...prev, botMessage]);

        } catch (err) {
            const errorMessage = { sender: 'bot' as const, text: 'Desculpe, não consegui processar sua solicitação no momento.' };
            setChatMessages(prev => [...prev, errorMessage]);
            console.error(err);
        } finally {
            setChatLoading(false);
        }
    };
    
    const openChat = () => {
        setChatOpen(true);
        if (chatMessages.length === 0) {
            if (chat) {
                setChatMessages([{ sender: 'bot', text: 'Olá! Como posso ajudar a analisar seus dados hoje?' }]);
            } else {
                setChatMessages([{ sender: 'bot', text: 'O assistente de IA não está disponível. A chave de API não foi configurada corretamente.' }]);
            }
        }
    };

    const handleResetApp = () => {
        setAllData([]);
        setFilters({ month: 'all', category: 'all', region: 'all' });
        setError('');
        setChat(null);
        setChatMessages([]);
        if (fileInputRef.current) {
            fileInputRef.current.value = "";
        }
    };

    return (
        <div className="app-layout">
            <Sidebar 
                theme={theme} 
                setTheme={setTheme} 
                isOpen={isSidebarOpen} 
                onUploadClick={() => fileInputRef.current?.click()}
                onResetApp={handleResetApp}
            />
            <input
                ref={fileInputRef}
                type="file"
                className="upload-input"
                multiple
                accept=".xlsx, .csv, .zip"
                onChange={e => handleFileUpload(e.target.files)}
            />
            {isSidebarOpen && <div className="sidebar-overlay" onClick={() => setSidebarOpen(false)}></div>}
            <div className="content-wrapper">
                <MainHeader 
                    onMenuClick={() => setSidebarOpen(!isSidebarOpen)} 
                    filters={filters}
                    filterOptions={filterOptions}
                    onFilterChange={handleFilterChange}
                    hasData={allData.length > 0}
                />
                <main className="main-content">
                    {isLoading && <div className="loader"></div>}
                    {error && <div className="error-message">{error}</div>}
                    {!isLoading && !error && !filteredData && (
                         <UploadArea 
                            onFileUpload={handleFileUpload} 
                            onAreaClick={() => fileInputRef.current?.click()} 
                        />
                    )}
                    {filteredData && <Dashboard data={filteredData} />}
                </main>
            </div>
            {filteredData && !isChatOpen && <ChatbotFab onOpen={openChat} />}
            <Chatbot 
                isOpen={isChatOpen} 
                onClose={() => setChatOpen(false)}
                messages={chatMessages}
                isLoading={isChatLoading}
                onSendMessage={handleSendMessage}
                suggestedQuestions={suggestedQuestions}
                isAvailable={!!chat}
            />
        </div>
    );
};

const Sidebar = ({ theme, setTheme, isOpen, onUploadClick, onResetApp }: { theme: string; setTheme: (theme: string) => void; isOpen: boolean; onUploadClick: () => void; onResetApp: () => void; }) => (
    <aside className={`sidebar ${isOpen ? 'open' : ''}`}>
        <div>
            <button className="sidebar-header" onClick={onResetApp}>
                <IconLogo />
                <h1>Alpha Insights</h1>
            </button>
            <nav>
                <a href="#" className="nav-item active">
                    <IconDashboard />
                    <span>Dashboard</span>
                </a>
                <button onClick={onUploadClick} className="nav-item">
                    <IconUpload />
                    <span>Upload Planilhas</span>
                </button>
            </nav>
        </div>
        <div className="theme-switcher">
            <span>Tema</span>
            <div>
                <button onClick={() => setTheme('light')} className={theme === 'light' ? 'active' : ''}><IconSun /></button>
                <button onClick={() => setTheme('dark')} className={theme === 'dark' ? 'active' : ''}><IconMoon /></button>
            </div>
        </div>
    </aside>
);

const MainHeader = ({ onMenuClick, filters, filterOptions, onFilterChange, hasData }: { 
    onMenuClick: () => void;
    filters: { month: string; category: string; region: string; };
    filterOptions: { months: string[]; categories: string[]; regions: string[]; };
    onFilterChange: (filterType: 'month' | 'category' | 'region', value: string) => void;
    hasData: boolean;
}) => (
    <header className="main-header">
        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
            <button className="hamburger" onClick={onMenuClick}><IconMenu /></button>
            <h2>Dashboard</h2>
        </div>
        {hasData && (
            <div className="filters">
                <div className="filter-group">
                    <select value={filters.month} onChange={e => onFilterChange('month', e.target.value)}>
                        <option value="all">Todos os Meses</option>
                        {filterOptions.months.map(m => <option key={m} value={m}>{new Date(m + '-02').toLocaleDateString('pt-BR', { year: 'numeric', month: 'long' })}</option>)}
                    </select>
                </div>
                <div className="filter-group">
                    <select value={filters.category} onChange={e => onFilterChange('category', e.target.value)}>
                        <option value="all">Todas as Categorias</option>
                        {filterOptions.categories.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                </div>
                <div className="filter-group">
                    <select value={filters.region} onChange={e => onFilterChange('region', e.target.value)}>
                        <option value="all">Todas as Regiões</option>
                         {filterOptions.regions.map(r => <option key={r} value={r}>{r}</option>)}
                    </select>
                </div>
            </div>
        )}
    </header>
);

const UploadArea = ({ onFileUpload, onAreaClick }: {
    onFileUpload: (files: FileList | null) => void;
    onAreaClick: () => void;
}) => (
    <section
        className="upload-section"
        onClick={onAreaClick}
        onDragOver={e => e.preventDefault()}
        onDrop={e => { e.preventDefault(); onFileUpload(e.dataTransfer.files); }}
    >
        <IconLogo />
        <h2>Olá! Seja bem-vindo(a) ao assistente de análise da Alpha Insights.</h2>
        <p className="welcome-subtitle">Sou o AI Data Analyst, seu bot especializado em transformar dados em respostas inteligentes.</p>
        <p className="upload-instruction">Para começar, arraste e solte seus arquivos (.xlsx, .csv ou .zip) aqui, ou clique para selecionar.</p>
    </section>
);

const Dashboard = ({ data }: { data: DashboardData }) => {
    const currencyFormatter = (value?: number) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value || 0);
    const numberFormatter = (value?: number) => new Intl.NumberFormat('pt-BR').format(value || 0);

    return (
        <>
            <section className="kpi-grid">
                <KPICard title="Receita Total" value={currencyFormatter(data.totalRevenue)} />
                <KPICard title="Ticket Médio" value={currencyFormatter(data.averageTicket)} />
                <KPICard title="Mês de Pico" value={data.bestMonth} />
            </section>
            <section className="charts-grid">
                <div className="chart-container" style={{ gridColumn: '1 / -1' }}><ChartWrapper type="bar" title="Vendas Mensais" chartData={data.monthlySalesChart} /></div>
                <div className="centered-chart-wrapper">
                    <div className="chart-container">
                        <ChartWrapper type="pie" title="Distribuição por Categoria" chartData={data.categorySalesChart} />
                    </div>
                </div>
            </section>
            <section className="table-container">
                <h3>Vendas Recentes</h3>
                <RecentSalesTable sales={data.recentSales} formatter={currencyFormatter} />
            </section>
        </>
    );
};

const KPICard = ({ title, value }: { title: string; value: string | number; }) => (
    <div className="kpi-card">
        <h3>{title}</h3>
        <div className="value">{value}</div>
    </div>
);

const RecentSalesTable = ({ sales, formatter }: { sales: CleanedSalesData[]; formatter: (value?: number) => string; }) => (
    <table className="sales-table">
        <thead>
            <tr>
                <th>Data</th>
                <th>ID Transação</th>
                <th>Produto</th>
                <th>Região</th>
                <th>Receita</th>
            </tr>
        </thead>
        <tbody>
            {sales.map((sale, index) => (
                <tr key={index}>
                    <td>{sale.Date.toLocaleDateString('pt-BR')}</td>
                    <td>{sale['Transaction ID'] || '-'}</td>
                    <td>{sale.Product}</td>
                    <td>{sale.Region}</td>
                    <td>{formatter(sale['Total Revenue'] || (sale.Quantity! * sale['Unit Price']!))}</td>
                </tr>
            ))}
        </tbody>
    </table>
);

const ChartWrapper = ({ type, title, chartData }: { type: 'bar' | 'pie' | 'doughnut' | 'line'; title: string; chartData: { labels: string[], data: number[] } }) => {
    const chartRef = useRef<HTMLCanvasElement>(null);
    const chartInstance = useRef<any | null>(null);
    const theme = document.body.getAttribute('data-theme');

    useEffect(() => {
        if (!chartRef.current || !chartData) return;
        
        const useDarkTheme = theme === 'dark';
        const textColor = useDarkTheme ? '#c9d1d9' : '#1e293b';
        const gridColor = useDarkTheme ? '#30363d' : '#e2e8f0';
        const borderColor = useDarkTheme ? '#0d1117' : '#ffffff';

        if (chartInstance.current) {
            chartInstance.current.destroy();
        }
        const ctx = chartRef.current.getContext('2d');
        if (!ctx) return;
        chartInstance.current = new Chart(ctx, {
            type: type,
            data: {
                labels: chartData.labels,
                datasets: [{
                    label: 'Vendas',
                    data: chartData.data,
                    backgroundColor: type === 'bar' ? 'rgba(0, 86, 179, 0.7)' : ['#0056b3', '#007bff', '#58a6ff', '#86bfff', '#b3d7ff'],
                    borderColor: borderColor,
                    borderWidth: type !== 'bar' ? 2 : 1,
                    tension: 0.3
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: false,
                plugins: {
                    legend: {
                        display: type !== 'bar',
                        position: 'top',
                        labels: { color: textColor }
                    },
                    title: {
                        display: true,
                        text: title,
                        color: textColor,
                        font: { size: 16, weight: '600' }
                    }
                },
                scales: ['bar', 'line'].includes(type) ? {
                    y: {
                        beginAtZero: true,
                        ticks: { color: textColor },
                        grid: { color: gridColor }
                    },
                    x: {
                        ticks: { color: textColor },
                        grid: { color: gridColor }
                    }
                } : {}
            }
        });

        return () => {
            if (chartInstance.current) {
                chartInstance.current.destroy();
            }
        };
    }, [chartData, type, title, theme]);

    return <canvas ref={chartRef}></canvas>;
};

const ChatbotFab = ({ onOpen }: { onOpen: () => void }) => (
    <button className="chatbot-fab" onClick={onOpen} aria-label="Abrir chat">
        <IconChat />
    </button>
);

const Chatbot = ({ isOpen, onClose, messages, isLoading, onSendMessage, suggestedQuestions = [], isAvailable }: {
    isOpen: boolean;
    onClose: () => void;
    messages: { sender: 'user' | 'bot'; text: string; }[];
    isLoading: boolean;
    onSendMessage: (message: string) => void;
    suggestedQuestions?: string[];
    isAvailable: boolean;
}) => {
    const [input, setInput] = useState('');
    const [showSuggestions, setShowSuggestions] = useState(true);
    const messagesEndRef = useRef<HTMLDivElement>(null);

    useEffect(() => {
        messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
    }, [messages]);
    
    const handleSend = (message: string) => {
        if (!message.trim() || isLoading || !isAvailable) return;
        onSendMessage(message);
        setInput('');
        setShowSuggestions(false);
    };

    const handleSubmit = (e: React.FormEvent) => {
        e.preventDefault();
        handleSend(input);
    };

    if (!isOpen) return null;

    return (
        <div className="chatbot-container">
            <header className="chatbot-header">
                <h3>Assistente Alpha</h3>
                <button onClick={onClose}>&times;</button>
            </header>
            <div className="chatbot-messages">
                {messages.map((msg, index) => (
                    <div key={index} className={`message ${msg.sender === 'user' ? 'user-message' : 'bot-message'}`}>
                        <div dangerouslySetInnerHTML={{ __html: marked.parse(msg.text) }} />
                    </div>
                ))}
                {isLoading && <div className="message bot-message loading"><span>.</span><span>.</span><span>.</span></div>}
                <div ref={messagesEndRef} />
            </div>

            {isAvailable && showSuggestions && !isLoading && suggestedQuestions.length > 0 && (
                <div className="chatbot-suggestions">
                    {suggestedQuestions.map((q, i) => (
                        <button key={i} className="suggestion-chip" onClick={() => handleSend(q)}>
                            {q}
                        </button>
                    ))}
                </div>
            )}

            <form className="chatbot-input-form" onSubmit={handleSubmit}>
                <input
                    type="text"
                    value={input}
                    onChange={e => setInput(e.target.value)}
                    placeholder={isAvailable ? "Pergunte sobre os dados..." : "Assistente indisponível"}
                    disabled={isLoading || !isAvailable}
                    aria-label="Digite sua pergunta"
                />
                <button type="submit" disabled={isLoading || !input.trim() || !isAvailable} aria-label="Enviar mensagem">
                    <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="22" y1="2" x2="11" y2="13"></line><polygon points="22 2 15 22 11 13 2 9 22 2"></polygon></svg>
                </button>
            </form>
        </div>
    );
};

const container = document.getElementById('root');
if (container) {
    createRoot(container).render(<App />);
}