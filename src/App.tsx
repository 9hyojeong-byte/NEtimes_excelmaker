import React, { useState, useRef, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileUp, 
  Download, 
  FileSpreadsheet, 
  Trash2, 
  CheckCircle2, 
  AlertCircle,
  FileText,
  Layers,
  Settings,
  History,
  LayoutDashboard,
  PlusCircle,
  User,
  ClipboardPaste,
  Type
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from './lib/utils';

interface ExcelData {
  [key: string]: any;
}

type InputMode = 'manual' | 'delivery' | 'file';

interface DeliveryRow {
  이름: string;
  휴대폰번호: string;
  주소: string;
  상품명: string;
}

export default function App() {
  const [inputMode, setInputMode] = useState<InputMode>('manual');
  const [manualText, setManualText] = useState('');
  const [file, setFile] = useState<File | null>(null);
  const [data, setData] = useState<ExcelData[] | null>(null);
  const [deliveryData, setDeliveryData] = useState<DeliveryRow[] | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isConverted, setIsConverted] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [activePreviewTab, setActivePreviewTab] = useState<string | null>(null);
  
  // 로컬스토리지에서 초기값 불러오기
  const [productIssueNumbers, setProductIssueNumbers] = useState<{[key: string]: string}>(() => {
    const saved = localStorage.getItem('productIssueNumbers');
    return saved ? JSON.parse(saved) : {};
  });
  const [draftIssueNumbers, setDraftIssueNumbers] = useState<{[key: string]: string}>(() => {
    const saved = localStorage.getItem('draftIssueNumbers');
    return saved ? JSON.parse(saved) : {};
  });
  
  const [customSheetNames, setCustomSheetNames] = useState<{[key: string]: string}>({});
  const fileInputRef = useRef<HTMLInputElement>(null);
  const tempIssueValuesRef = useRef<{[key: string]: string}>({});

  // 변경될 때마다 로컬스토리지 저장
  React.useEffect(() => {
    localStorage.setItem('productIssueNumbers', JSON.stringify(productIssueNumbers));
  }, [productIssueNumbers]);

  React.useEffect(() => {
    localStorage.setItem('draftIssueNumbers', JSON.stringify(draftIssueNumbers));
  }, [draftIssueNumbers]);

  const applyIssueNumbers = () => {
    setProductIssueNumbers({ ...draftIssueNumbers });
  };

  const PRODUCT_NAME_MAPPINGS: {[key: string]: string} = {
    'NETIMES': 'times',
    'NETIMES JUNIOR 주간지': 'junior',
    'NETIMES KIDS': 'kids',
    'NETIMES KINDER': 'kinder'
  };

  const SORT_ORDER = ['times', 'junior', 'kids', 'kinder'];

  const getTransformedName = useCallback((originalName: string) => {
    const trimmed = originalName.trim();
    const base = PRODUCT_NAME_MAPPINGS[trimmed] || trimmed;
    const custom = customSheetNames[trimmed];
    if (custom) return custom;
    
    // 개별 호수 숫자 적용 (기본값 0)
    const issueNum = productIssueNumbers[trimmed] || '0';
    return `${base} ${issueNum}호`;
  }, [productIssueNumbers, customSheetNames]);

  const REQUIRED_COLUMNS = ['발송구분', '이름', '우편번호', '주소', '상품명', '수량', '구독기간', '전달사항'];
  const EXPORT_COLUMNS = ['이름', '우편번호', '주소', '상품명', '구독기간'];

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFile = e.target.files?.[0];
    if (uploadedFile) {
      processFile(uploadedFile);
    }
  };

  const processFile = (file: File) => {
    setFile(file);
    setError(null);
    setIsConverted(false);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const ab = e.target?.result;
        if (!ab) throw new Error('파일 데이터를 읽지 못했습니다.');

        const wb = XLSX.read(ab, { type: 'array', cellDates: true, cellNF: false, cellText: false });
        
        if (!wb.SheetNames || wb.SheetNames.length === 0) {
          throw new Error('엑셀 파일에 유효한 시트가 없습니다.');
        }

        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        
        const rawJsonData = XLSX.utils.sheet_to_json(ws, { defval: "" }) as any[];
        handleDataLoad(rawJsonData);
        
      } catch (err: any) {
        console.error('Excel Process Error:', err);
        setError(err.message || '파일을 분석하는 중 알 수 없는 오류가 발생했습니다.');
        setData(null);
        setFile(null);
      }
    };
    reader.onerror = () => {
      setError('파일을 읽는 과정에서 시스템 오류가 발생했습니다.');
    };
    reader.readAsArrayBuffer(file);
  };

  const handleManualSubmit = () => {
    if (!manualText.trim()) {
      setError('데이터를 입력해주세요.');
      return;
    }

    try {
      setError(null);
      const rows = manualText.trim().split('\n');
      if (rows.length < 2) {
        throw new Error('적어도 제목행과 한 개 이상의 데이터 행이 필요합니다.');
      }

      // 탭(\t) 또는 쉼표(,) 구분 감지
      const firstRow = rows[0];
      const separator = firstRow.includes('\t') ? '\t' : (firstRow.includes(',') ? ',' : ' ');
      
      const headers = rows[0].split(separator).map(h => h.trim());
      const jsonData = rows.slice(1).map(rowStr => {
        const values = rowStr.split(separator);
        const row: any = {};
        headers.forEach((header, index) => {
          row[header] = values[index] ? values[index].trim() : "";
        });
        return row;
      });

      handleDataLoad(jsonData);
      setFile(new File([], "직접입력_데이터.txt")); 
    } catch (err: any) {
      setError('데이터를 분석하는 중 오류가 발생했습니다. 데이터를 정확히 복사했는지 확인해주세요.');
    }
  };

  const handleDataLoad = (rawJsonData: any[]) => {
    if (rawJsonData.length === 0) {
      throw new Error('데이터가 비어있거나 올바른 형식이 아닙니다.');
    }

    const jsonData = rawJsonData.map((row: any) => {
      const normalizedRow: any = {};
      Object.keys(row).forEach(key => {
        const cleanKey = String(key).trim();
        normalizedRow[cleanKey] = row[key];
      });
      return normalizedRow;
    });

    setData(jsonData);
    
    const sampleRow = jsonData[0];
    const actualKeys = Object.keys(sampleRow);
    const productNameKey = actualKeys.find(k => k.replace(/\s/g, '') === '상품명') || '상품명';

    const firstProduct = String(jsonData[0][productNameKey] || '기타');
    setActivePreviewTab(firstProduct);
  };

  const groupedPreview = useMemo(() => {
    if (!data) return {};
    const grouped: { [key: string]: ExcelData[] } = {};
    data.forEach(row => {
      const keys = Object.keys(row);
      const pKey = keys.find(k => k.replace(/\s/g, '') === '상품명') || '상품명';
      const productName = String(row[pKey] || '기타').trim();
      const transformedName = getTransformedName(productName);
      if (!grouped[transformedName]) grouped[transformedName] = [];
      grouped[transformedName].push(row);
    });
    return grouped;
  }, [data, getTransformedName]);

  const productNames = useMemo(() => {
    const keys = Object.keys(groupedPreview);
    return keys.sort((a, b) => {
      // 변환된 이름에서 base 추출 (공백 기준 첫 단어)
      const getBase = (name: string) => {
        const foundOrig = Object.keys(PRODUCT_NAME_MAPPINGS).find(orig => getTransformedName(orig) === name);
        return foundOrig ? PRODUCT_NAME_MAPPINGS[foundOrig] : name.split(' ')[0];
      };
      
      const orderA = SORT_ORDER.indexOf(getBase(a));
      const orderB = SORT_ORDER.indexOf(getBase(b));
      
      const valA = orderA === -1 ? 999 : orderA;
      const valB = orderB === -1 ? 999 : orderB;
      
      return valA - valB;
    });
  }, [groupedPreview, getTransformedName]);

  const convertData = useCallback(() => {
    const isDelivery = inputMode === 'delivery';
    if (isDelivery) {
      if (!deliveryData || deliveryData.length === 0) return;
    } else {
      if (!data || !productNames.length) return;
    }

    setIsProcessing(true);
    setTimeout(() => {
      try {
        const newWb = XLSX.utils.book_new();

        if (isDelivery && deliveryData) {
          const exportCols = ['이름', '휴대폰번호', '주소', '상품명'];
          const ws = XLSX.utils.json_to_sheet(deliveryData, { header: exportCols });
          XLSX.utils.book_append_sheet(newWb, ws, "택배발송리스트");
        } else if (data) {
          // 미리보기 탭 순서(productNames)대로 시트 생성
          productNames.forEach(transformedName => {
            const rows = groupedPreview[transformedName];
            if (!rows || rows.length === 0) return;

            // 미리보기에 보이는 데이터와 동일하게 EXPORT_COLUMNS 기준으로 매핑
            const formattedRows = rows.map(row => {
              const keys = Object.keys(row);
              const filteredRow: any = {};
              
              EXPORT_COLUMNS.forEach(col => {
                const targetKey = keys.find(k => k.replace(/\s/g, '') === col.replace(/\s/g, '')) || col;
                let value = row[targetKey] || '';
                
                // 상품명 컬럼의 경우 현재 변환된 이름(탭 이름)으로 고정
                if (col === '상품명') {
                  value = transformedName;
                }
                
                filteredRow[col] = value;
              });
              return filteredRow;
            });

            const safeSheetName = transformedName.substring(0, 31).replace(/[\\/?*[\]]/g, '');
            const ws = XLSX.utils.json_to_sheet(formattedRows, { header: EXPORT_COLUMNS });
            XLSX.utils.book_append_sheet(newWb, ws, safeSheetName);
          });
        }

        const now = new Date();
        const yy = String(now.getFullYear()).slice(2);
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const dd = String(now.getDate()).padStart(2, '0');
        const prefix = isDelivery ? '택배발송리스트' : '정기발송리스트';
        const fileName = `${prefix}_${yy}${mm}${dd}.xlsx`;

        XLSX.writeFile(newWb, fileName);
        setIsConverted(true);
      } catch (err: any) {
        console.error('Export Error:', err);
        setError('변환 중 오류가 발생했습니다.');
      } finally {
        setIsProcessing(false);
      }
    }, 1200);
  }, [data, deliveryData, file, productNames, groupedPreview, inputMode]);

  const reset = () => {
    setFile(null);
    setData(null);
    setDeliveryData(null);
    setIsConverted(false);
    setError(null);
    setActivePreviewTab(null);
    setManualText('');
    setCustomSheetNames({});
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleDeliverySubmit = (overrideIssueNumbers?: {[key: string]: string}) => {
    if (!manualText.trim()) {
      setError('데이터를 입력해주세요.');
      return;
    }

    const currentIssues = overrideIssueNumbers || productIssueNumbers;

    try {
      setError(null);
      const rows = manualText.trim().split('\n').filter(r => r.trim());
      if (rows.length < 2) {
        throw new Error('적어도 제목행과 한 개 이상의 데이터 행이 필요합니다.');
      }

      const firstRow = rows[0];
      const separator = firstRow.includes('\t') ? '\t' : (firstRow.includes(',') ? ',' : ' ');
      const headers = rows[0].split(separator).map(h => h.trim());
      
      const rawRows = rows.slice(1).map(rowStr => {
        const values = rowStr.split(separator);
        const row: any = {};
        headers.forEach((header, index) => {
          // 헤더에서 공백 보이지 않는 문자 제거 후 소문자로 저장하되, 원래 키도 매핑
          const cleanHeader = header.replace(/\s/g, '');
          row[cleanHeader] = values[index] ? values[index].trim() : "";
        });
        return row;
      });

      // 주소 기준 그룹화 (트림 처리 추가)
      const groupedByAddress: { [address: string]: any[] } = {};
      rawRows.forEach(row => {
        const addrKey = Object.keys(row).find(k => k === '주소') || '주소';
        const addr = (row[addrKey] || '').trim();
        if (!addr) return; 
        if (!groupedByAddress[addr]) groupedByAddress[addr] = [];
        groupedByAddress[addr].push(row);
      });

      const processedDelivery: DeliveryRow[] = Object.entries(groupedByAddress).map(([address, group]) => {
        // 1. 이름 및 휴대폰번호 결정
        let finalName = '';
        let finalPhone = '';

        const nameInAddress = group.find(item => {
          const nameKey = Object.keys(item).find(k => k === '이름') || '이름';
          const name = String(item[nameKey] || '').replace('선생님', '').trim();
          return name && address.includes(name);
        });

        if (nameInAddress) {
          const nameKey = Object.keys(nameInAddress).find(k => k === '이름') || '이름';
          const phoneKey = Object.keys(nameInAddress).find(k => k === '휴대폰번호' || k === '휴대폰') || '휴대폰번호';
          finalName = nameInAddress[nameKey];
          finalPhone = nameInAddress[phoneKey];
        } else {
          const nameKey = Object.keys(group[0]).find(k => k === '이름') || '이름';
          const phoneKey = Object.keys(group[0]).find(k => k === '휴대폰번호' || k === '휴대폰') || '휴대폰번호';
          finalName = group[0][nameKey];
          finalPhone = group[0][phoneKey];
        }

        // 2. 상품명 가공
        const brandsCount: { [brand: string]: number } = {};
        let totalCount = 0;

        const productMappings = Object.keys(PRODUCT_NAME_MAPPINGS).sort((a, b) => b.length - a.length);

        group.forEach(item => {
          // 상품명 키 찾기 (공백 제거 후 비교)
          const pKey = Object.keys(item).find(k => k.replace(/\s/g, '') === '상품명') || '상품명';
          const rawProductName = (item[pKey] || '').trim();

          // 한 셀에 여러 상품이 쉼표나 세미콜론으로 구분되어 있는 경우 대응
          const parts = rawProductName.split(/[,,;]/).map(p => p.trim()).filter(p => p);
          
          let rowHandled = false;
          parts.forEach(part => {
            // 대소문자 무시 매칭 (가장 구체적인 이름부터)
            const brandKey = productMappings.find(k => part.toUpperCase().includes(k.toUpperCase()));
            
            if (brandKey) {
              rowHandled = true;
              const mappedName = PRODUCT_NAME_MAPPINGS[brandKey];
              
              // 부수 추출 로직 (1순위: 파트 내 'n개' 패턴, 2순위: 별도 수량 컬럼)
              const 개Match = part.match(/(\d+)\s*개/);
              let count = 0;
              
              if (개Match) {
                count = parseInt(개Match[1]);
              } else {
                const cKey = Object.keys(item).find(k => ['총부수', '부수', '수량', '합계'].includes(k.replace(/\s/g, ''))) || '총부수';
                const rawCountValue = item[cKey] || '1';
                count = parseInt(String(rawCountValue).replace(/[^0-9]/g, '')) || 0;
              }
              
              brandsCount[mappedName] = (brandsCount[mappedName] || 0) + count;
              totalCount += count;
            }
          });

          // 매칭되는 브랜드가 전혀 없는 경우
          if (!rowHandled && rawProductName) {
            const cKey = Object.keys(item).find(k => ['총부수', '부수', '수량', '합계'].includes(k.replace(/\s/g, ''))) || '총부수';
            const rawCountValue = item[cKey] || '1';
            const count = parseInt(String(rawCountValue).replace(/[^0-9]/g, '')) || 0;
            brandsCount[rawProductName] = (brandsCount[rawProductName] || 0) + count;
            totalCount += count;
          }
        });

        const productParts = Object.entries(brandsCount)
          .sort((a, b) => {
            const orderA = SORT_ORDER.indexOf(a[0]);
            const orderB = SORT_ORDER.indexOf(b[0]);
            return (orderA === -1 ? 999 : orderA) - (orderB === -1 ? 999 : orderB);
          })
          .map(([brand, count]) => {
            // 현재 적용된 호수 가져오기
            const origKey = Object.keys(PRODUCT_NAME_MAPPINGS).find(k => PRODUCT_NAME_MAPPINGS[k] === brand) || brand;
            const issueNum = currentIssues[origKey] || '0';
            return `${brand} ${issueNum}호 ${count}부`;
          });

        const finalProductString = `${productParts.join(', ')} (총${totalCount}부)`;

        return {
          이름: finalName,
          휴대폰번호: finalPhone,
          주소: address,
          상품명: finalProductString
        };
      });

      setDeliveryData(processedDelivery);
      setFile(new File([], "택배발송리스트_데이터.txt"));
      setIsConverted(false);
    } catch (err: any) {
      console.error(err);
      setError('데이터 분석 중 오류가 발생했습니다. 형식을 확인해주세요.');
    }
  };

  // activePreviewTab 동기화: 만약 현재 탭이 결과에 없다면 첫 번째 탭으로 이동
  React.useEffect(() => {
    if (productNames.length > 0) {
      if (!activePreviewTab || !productNames.includes(activePreviewTab)) {
        setActivePreviewTab(productNames[0]);
      }
    } else {
      setActivePreviewTab(null);
    }
  }, [productNames, activePreviewTab]);

  return (
    <div className="flex h-screen w-full bg-[#F8FAFC] text-[#1E293B] font-sans overflow-hidden">
      <aside className="w-60 bg-[#0F172A] text-[#F8FAFC] flex flex-col flex-shrink-0">
        <div className="p-6 border-b border-[#1E293B]">
          <div className="text-lg font-bold tracking-tight flex items-center">
            <div className="w-8 h-8 bg-[#3B82F6] rounded-lg flex items-center justify-center mr-3">
              <Layers size={18} />
            </div>
            NEtimes 발송요청 엑셀변환기
          </div>
        </div>
        
        <nav className="py-4 flex-grow overflow-y-auto">
          <div className="px-6 py-3 text-[11px] font-bold uppercase tracking-wider text-[#64748B] mb-1">Transformers</div>
          
          <div 
            onClick={() => { setInputMode('manual'); reset(); }}
            className={cn(
              "flex items-center px-6 py-3 text-sm transition-all cursor-pointer",
              inputMode === 'manual' 
                ? "opacity-100 bg-[#1E293B] border-r-4 border-[#3B82F6]" 
                : "opacity-60 hover:opacity-100"
            )}
          >
            <ClipboardPaste size={16} className="mr-3" />
            정기발송리스트 (복붙)
          </div>

          <div 
            onClick={() => { setInputMode('delivery'); reset(); }}
            className={cn(
              "flex items-center px-6 py-3 text-sm transition-all cursor-pointer",
              inputMode === 'delivery' 
                ? "opacity-100 bg-[#1E293B] border-r-4 border-[#3B82F6]" 
                : "opacity-60 hover:opacity-100"
            )}
          >
            <FileText size={16} className="mr-3" />
            택배발송리스트 (복붙)
          </div>
        </nav>

        <div className="p-6 border-t border-[#1E293B] text-[11px] text-[#64748B]">
          <div className="font-bold text-[#F8FAFC] mb-1">v1.3.0 Professional</div>
          © 2024 AI Systems Inc.
        </div>
      </aside>

      <main className="flex-grow flex flex-col h-full overflow-hidden">
        <header className="h-16 bg-white border-b border-[#E2E8F0] flex items-center justify-between px-8 flex-shrink-0">
          <div className="font-semibold text-base">
            {inputMode === 'delivery' ? '택배발송리스트 데이터 합치기' : '정기발송리스트 데이터 가공'}
          </div>
          <div className="flex items-center gap-6">
            <div className="flex items-center text-xs text-[#64748B] gap-2">
              <User size={14} />
              <span>ID: user_8293</span>
            </div>
            <button 
              onClick={reset}
              className="bg-[#3B82F6] hover:bg-[#2563EB] text-white px-4 py-2 rounded-md font-semibold text-sm transition-colors flex items-center gap-2"
            >
              <PlusCircle size={16} />
              새 작업 시작
            </button>
          </div>
        </header>

        <div className="flex-grow flex flex-col overflow-hidden p-6 space-y-6">
          <section className="bg-white rounded-xl border border-[#E2E8F0] shadow-sm p-6 flex-shrink-0">
            {!data && !deliveryData ? (
              <div className="space-y-4">
                <div className="text-sm font-semibold text-[#475569] flex items-center gap-2">
                  <ClipboardPaste size={16} />
                  엑셀의 데이터를 아래에 붙여넣으세요 (제목행 포함)
                </div>
                <textarea
                  value={manualText}
                  onChange={(e) => setManualText(e.target.value)}
                  placeholder={inputMode === 'delivery' 
                    ? "발송구분, 이름, 휴대폰번호, 주소, 상품명, 총 부수... 영역을 붙여넣으세요"
                    : "상품명, 이름, 주소 등의 데이터 영역을 드래깅하여 여기에 붙여넣으세요..."}
                  className="w-full h-40 p-4 border border-[#CBD5E1] rounded-lg bg-[#F1F5F9] text-sm font-mono focus:outline-none focus:ring-2 focus:ring-[#3B82F6]/20 transition-all resize-none"
                />

                {/* Relocated Issue Number UI */}
                <div className="bg-[#F8FAFC] border border-[#E2E8F0] p-4 rounded-lg">
                  <div className="text-xs font-bold text-[#64748B] uppercase tracking-wider mb-3 flex items-center gap-2">
                    <PlusCircle size={14} />
                    매체별 호수 지정 (가공 시 반영됨)
                  </div>
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                    {Object.keys(PRODUCT_NAME_MAPPINGS).sort((a, b) => {
                      const orderA = SORT_ORDER.indexOf(PRODUCT_NAME_MAPPINGS[a]);
                      const orderB = SORT_ORDER.indexOf(PRODUCT_NAME_MAPPINGS[b]);
                      return orderA - orderB;
                    }).map(orig => (
                      <div key={orig} className="flex flex-col gap-1.5">
                        <label className="text-[10px] font-bold text-[#94A3B8] uppercase">{PRODUCT_NAME_MAPPINGS[orig]}</label>
                        <div className="flex items-center bg-white border border-[#CBD5E1] rounded-md px-2 focus-within:ring-2 focus-within:ring-blue-500/20 transition-all shadow-sm">
                          <input 
                            type="text"
                            inputMode="numeric"
                            value={draftIssueNumbers[orig] || ''}
                            placeholder="0"
                            onFocus={(e) => {
                              tempIssueValuesRef.current[orig] = e.target.value;
                              setDraftIssueNumbers(prev => ({ ...prev, [orig]: '' }));
                            }}
                            onBlur={(e) => {
                              if (e.target.value === '') {
                                setDraftIssueNumbers(prev => ({ ...prev, [orig]: tempIssueValuesRef.current[orig] || '' }));
                              }
                            }}
                            onChange={(e) => {
                              const val = e.target.value.replace(/[^0-9]/g, '');
                              setDraftIssueNumbers(prev => ({ ...prev, [orig]: val }));
                            }}
                            className="w-full py-1.5 text-xs font-bold text-[#1E293B] outline-none text-right pr-1"
                          />
                          <span className="text-[11px] text-[#64748B] font-medium ml-1">호</span>
                        </div>
                      </div>
                    ))}
                  </div>
                  <div className="mt-3 flex justify-end">
                    <button
                      onClick={() => {
                        applyIssueNumbers();
                      }}
                      className="bg-[#1E293B] hover:bg-black text-white text-[11px] font-bold px-4 py-2 rounded shadow-sm transition-colors flex items-center gap-2"
                    >
                      <CheckCircle2 size={13} />
                      호수 설정 저장
                    </button>
                  </div>
                </div>

                <button
                  onClick={() => {
                    if (inputMode === 'delivery') {
                      applyIssueNumbers();
                      handleDeliverySubmit(draftIssueNumbers);
                    } else {
                      applyIssueNumbers(); // Ensure latest issues are used for manual too
                      handleManualSubmit();
                    }
                  }}
                  className="w-full py-3 bg-[#3B82F6] text-white rounded-lg font-bold text-sm hover:bg-[#2563EB] transition-colors shadow-lg shadow-blue-500/20 active:scale-[0.99]"
                >
                  데이터 변환 시작
                </button>
              </div>
            ) : (
              <div className="relative group rounded-lg p-6 bg-[#F1F5F9] border border-[#CBD5E1]">
                <div className="flex items-center justify-between">
                  <div>
                    <div className="text-sm font-bold text-[#475569] flex items-center gap-2">
                      <CheckCircle2 size={16} className="text-emerald-500" />
                      직접 입력 데이터 분석 완료
                    </div>
                    <div className="text-[11px] text-[#94A3B8] mt-1">
                      {(data?.length || deliveryData?.length)}개의 데이터 행이 인식되었습니다.
                    </div>
                  </div>
                  <button 
                    onClick={reset}
                    className="p-1 px-2 text-[10px] bg-white border border-[#E2E8F0] text-[#64748B] rounded hover:bg-white hover:text-red-500 transition-colors"
                  >
                    초기화 후 다시 입력
                  </button>
                </div>
              </div>
            )}
          </section>

          <section className={cn(
            "flex-grow bg-white rounded-xl border border-[#E2E8F0] flex flex-col overflow-hidden",
            (!data && !deliveryData) && "opacity-40 pointer-events-none"
          )}>
            <div className="flex px-4 pt-4 border-b border-[#E2E8F0] bg-[#F8FAFC] gap-1 shrink-0 overflow-x-auto no-scrollbar">
              {inputMode === 'delivery' ? (
                <button className="px-4 py-2 font-semibold text-[13px] rounded-t-md bg-white text-[#3B82F6] border-t border-x border-[#E2E8F0] -mb-[1px]">
                  전체 미리보기 ({deliveryData?.length || 0})
                </button>
              ) : (
                data ? productNames.map(name => (
                  <button
                    key={name}
                    onClick={() => setActivePreviewTab(name)}
                    className={cn(
                      "px-4 py-2 font-semibold text-[13px] rounded-t-md transition-all whitespace-nowrap outline-none",
                      activePreviewTab === name 
                        ? "bg-white text-[#3B82F6] border-t border-x border-[#E2E8F0] -mb-[1px]" 
                        : "text-[#64748B] hover:text-[#1E293B]"
                    )}
                  >
                    {name} ({groupedPreview[name].length})
                  </button>
                )) : (
                  <div className="px-4 py-2 text-[13px] font-semibold text-[#64748B]">샘플 탭 (0)</div>
                )
              )}
            </div>

            <div className="flex-grow overflow-auto bg-white">
              <table className="w-full border-collapse text-[13px]">
                <thead className="sticky top-0 bg-[#F8FAFC] z-10">
                  <tr>
                    {inputMode === 'delivery' ? (
                      ['이름', '휴대폰번호', '주소', '상품명'].map(col => (
                        <th key={col} className="text-left px-4 py-3 font-semibold text-[#475569] border-bottom border-[#E2E8F0]">
                          {col}
                        </th>
                      ))
                    ) : (
                      EXPORT_COLUMNS.map(col => (
                        <th key={col} className="text-left px-4 py-3 font-semibold text-[#475569] border-bottom border-[#E2E8F0]">
                          {col}
                        </th>
                      ))
                    )}
                    <th className="text-left px-4 py-3 font-semibold text-[#475569] border-bottom border-[#E2E8F0]">상태</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#F1F5F9]">
                  {inputMode === 'delivery' ? (
                    deliveryData ? deliveryData.map((row, idx) => (
                      <tr key={idx} className="hover:bg-[#F8FAFC] transition-colors">
                        <td className="px-4 py-4 text-[#334155]">{row.이름}</td>
                        <td className="px-4 py-4 text-[#334155]">{row.휴대폰번호}</td>
                        <td className="px-4 py-4 text-[#334155]">{row.주소}</td>
                        <td className="px-4 py-4 text-[#334155] font-medium">{row.상품명}</td>
                        <td className="px-4 py-4">
                          <span className="w-2 h-2 rounded-full bg-blue-500 inline-block"></span>
                        </td>
                      </tr>
                    )) : (
                      Array.from({ length: 5 }).map((_, i) => (
                        <tr key={i} className="opacity-10">
                          {['이름', '휴대폰번호', '주소', '상품명'].map(col => (
                            <td key={col} className="px-4 py-4"><div className="h-4 bg-gray-200 rounded animate-pulse"></div></td>
                          ))}
                          <td className="px-4 py-4"><div className="h-4 w-12 bg-gray-200 rounded animate-pulse"></div></td>
                        </tr>
                      ))
                    )
                  ) : (
                    data && activePreviewTab && groupedPreview[activePreviewTab] ? (
                      groupedPreview[activePreviewTab].map((row, idx) => {
                        const keys = Object.keys(row);
                        return (
                          <tr key={idx} className="hover:bg-[#F8FAFC] transition-colors">
                            {EXPORT_COLUMNS.map(col => {
                              const targetKey = keys.find(k => k.replace(/\s/g, '') === col.replace(/\s/g, '')) || col;
                              let cellValue = row[targetKey] || '-';
                              
                              if (col === '상품명') {
                                cellValue = getTransformedName(String(cellValue));
                              }
                              
                              return <td key={col} className="px-4 py-4 text-[#334155]">{cellValue}</td>;
                            })}
                            <td className="px-4 py-4">
                              <span className="w-2 h-2 rounded-full bg-emerald-500 inline-block"></span>
                            </td>
                          </tr>
                        );
                      })
                    ) : (
                      Array.from({ length: 5 }).map((_, i) => (
                        <tr key={i} className="opacity-10">
                          {EXPORT_COLUMNS.map(col => (
                            <td key={col} className="px-4 py-4"><div className="h-4 bg-gray-200 rounded animate-pulse"></div></td>
                          ))}
                          <td className="px-4 py-4"><div className="h-4 w-12 bg-gray-200 rounded animate-pulse"></div></td>
                        </tr>
                      ))
                    )
                  )}
                </tbody>
              </table>
            </div>

            <div className="p-4 px-8 border-t border-[#E2E8F0] flex justify-between items-center bg-[#F8FAFC] shrink-0">
              <div className="text-xs text-[#64748B] flex items-center gap-2">
                {data || deliveryData ? (
                  <>
                    <span className="font-bold text-[#1E293B]">분석 완료:</span> 
                    총 {(data || deliveryData)?.length}행 로드됨
                  </>
                ) : (
                  "데이터를 입력하면 미리보기를 확인하실 수 있습니다."
                )}
              </div>
              <div className="flex items-center gap-3">
                {isConverted && (
                  <span className="flex items-center gap-1 text-[11px] font-bold text-emerald-600 mr-4">
                    <CheckCircle2 size={14} />
                    다운로드 완료
                  </span>
                )}
                <button 
                  disabled={(!data && !deliveryData) || isProcessing}
                  onClick={convertData}
                  className={cn(
                    "px-6 py-2 rounded-md font-bold text-[13px] transition-all flex items-center gap-2 shadow-sm",
                    (data || deliveryData) && !isProcessing
                      ? "bg-[#059669] hover:bg-[#047857] text-white cursor-pointer active:scale-95"
                      : "bg-[#CBD5E1] text-white cursor-not-allowed"
                  )}
                >
                  {isProcessing ? (
                    <>
                      <div className="w-3 h-3 border-2 border-white/40 border-t-white rounded-full animate-spin" />
                      처리 중...
                    </>
                  ) : (
                    <>
                      <Download size={16} />
                      Excel 파일 다운로드 (.xlsx)
                    </>
                  )}
                </button>
              </div>
            </div>
          </section>
        </div>

        {error && (
          <div className="fixed bottom-8 right-8 z-[100]">
            <motion.div 
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              className="bg-white border-l-4 border-red-500 shadow-2xl p-4 rounded-lg flex items-start gap-4 max-w-sm"
            >
              <AlertCircle className="text-red-500 shrink-0" size={20} />
              <div>
                <div className="font-bold text-sm text-[#1E293B]">데이터 처리 오류</div>
                <div className="text-xs text-[#64748B] mt-1">{error}</div>
                <button 
                  onClick={() => setError(null)}
                  className="text-[10px] font-bold text-red-500 mt-2 uppercase underline"
                >
                  닫기
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </main>
    </div>
  );
}
