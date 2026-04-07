import { useState, useRef, useEffect } from 'react';
import { GoogleGenAI, GenerateContentResponse, ThinkingLevel } from "@google/genai";
import { MessageSquare, X, Send, Loader2, Search, Bot, User, Minimize2, Maximize2 } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import ReactMarkdown from 'react-markdown';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface AIChatBotProps {
  data: {
    employees: any[];
    sales: any[];
    incentives: any[];
    incentivePayments: any[];
    bdes: any[];
    weeks: any[];
  };
  currentUser: any;
}

export const AIChatBot = ({ data, currentUser }: AIChatBotProps) => {
  const [isOpen, setIsOpen] = useState(false);
  const [isMinimized, setIsMinimized] = useState(false);
  const [messages, setMessages] = useState<{ role: 'user' | 'model'; text: string; sources?: any[] }[]>([]);
  const [input, setInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [messages, isLoading]);

  const handleSend = async () => {
    if (!input.trim() || isLoading) return;

    const userMessage = input.trim();
    setInput('');
    setMessages(prev => [...prev, { role: 'user', text: userMessage }]);
    setIsLoading(true);

    try {
      const ai = new GoogleGenAI({ apiKey: (process.env.GEMINI_API_KEY as string) });
      
      // Prepare data summary for context
      const dataSummary = `
        Current User: ${currentUser?.name || 'Unknown'} (${currentUser?.email || 'No email'})
        
        Application Data Summary:
        - Total Employees: ${data.employees.length}
        - Total Sales: ${data.sales.length}
        - Total Incentives: ${data.incentives.length}
        - Total Incentive Payments: ${data.incentivePayments.length}
        - Total BDEs: ${data.bdes.length}
        
        Context: You are an AI assistant for Nitsa Bravo, a sales management platform. 
        You have access to the application's data and can also search the web for external information.
        When answering questions about the data, be precise. If you don't have specific details, say so.
        Always respect user privacy and only share information the user should have access to.
        
        Data Details (JSON):
        ${JSON.stringify({
          employees: data.employees.map(e => ({ id: e.id, name: e.name, position: e.position, status: e.status })),
          salesSummary: data.sales.slice(-20).map(s => ({ date: s.date, amount: s.packageValue, bde: s.bde, status: s.status })),
          incentivesSummary: data.incentives.slice(-20).map(i => ({ date: i.date, amount: i.amount, type: i.type, employeeId: i.employeeId }))
        })}
      `;

      const response = await ai.models.generateContent({
        model: "gemini-3.1-flash-lite-preview",
        contents: [
          ...messages.map(m => ({
            role: m.role,
            parts: [{ text: m.text }]
          })),
          { role: 'user', parts: [{ text: userMessage }] }
        ],
        config: {
          systemInstruction: dataSummary,
          tools: [{ googleSearch: {} }],
          thinkingConfig: { thinkingLevel: ThinkingLevel.MINIMAL }
        },
      });

      const text = response.text || "I'm sorry, I couldn't generate a response.";
      const sources = response.candidates?.[0]?.groundingMetadata?.groundingChunks;

      setMessages(prev => [...prev, { role: 'model', text, sources }]);
    } catch (error) {
      console.error("ChatBot Error:", error);
      setMessages(prev => [...prev, { role: 'model', text: "Sorry, I encountered an error. Please try again." }]);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <>
      {/* Floating Action Button */}
      <button
        onClick={() => setIsOpen(true)}
        className="fixed bottom-6 right-6 w-14 h-14 bg-orange-600 text-white rounded-full shadow-lg flex items-center justify-center hover:bg-orange-700 transition-all z-50 group"
      >
        <MessageSquare className="w-6 h-6 group-hover:scale-110 transition-transform" />
      </button>

      <AnimatePresence>
        {isOpen && (
          <motion.div
            initial={{ opacity: 0, y: 20, scale: 0.95 }}
            animate={{ 
              opacity: 1, 
              y: 0, 
              scale: 1,
              height: isMinimized ? '64px' : '600px',
              width: isMinimized ? '300px' : '400px'
            }}
            exit={{ opacity: 0, y: 20, scale: 0.95 }}
            className="fixed bottom-6 right-6 bg-white rounded-2xl shadow-2xl border border-zinc-200 flex flex-col overflow-hidden z-50 max-w-[calc(100vw-3rem)]"
          >
            {/* Header */}
            <div className="bg-orange-600 p-4 text-white flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2">
                <Bot className="w-5 h-5" />
                <span className="font-bold text-sm uppercase tracking-wider">Nitsa AI Assistant</span>
              </div>
              <div className="flex items-center gap-1">
                <button 
                  onClick={() => setIsMinimized(!isMinimized)}
                  className="p-1.5 hover:bg-white/10 rounded-lg transition-colors"
                >
                  {isMinimized ? <Maximize2 className="w-4 h-4" /> : <Minimize2 className="w-4 h-4" />}
                </button>
                <button 
                  onClick={() => setIsOpen(false)}
                  className="p-1.5 hover:bg-white/10 rounded-lg transition-colors"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
            </div>

            {!isMinimized && (
              <>
                {/* Messages */}
                <div 
                  ref={scrollRef}
                  className="flex-1 overflow-y-auto p-4 space-y-4 bg-zinc-50"
                >
                  {messages.length === 0 && (
                    <div className="h-full flex flex-col items-center justify-center text-center p-8 space-y-4">
                      <div className="w-12 h-12 bg-orange-100 text-orange-600 rounded-full flex items-center justify-center">
                        <Bot className="w-6 h-6" />
                      </div>
                      <div>
                        <h4 className="font-bold text-zinc-900">How can I help you today?</h4>
                        <p className="text-xs text-zinc-500 mt-1">
                          I can analyze sales data, check incentive statuses, or search for information on Google.
                        </p>
                      </div>
                    </div>
                  )}
                  {messages.map((m, i) => (
                    <div 
                      key={i}
                      className={cn(
                        "flex gap-3",
                        m.role === 'user' ? "flex-row-reverse" : "flex-row"
                      )}
                    >
                      <div className={cn(
                        "w-8 h-8 rounded-full flex items-center justify-center shrink-0",
                        m.role === 'user' ? "bg-zinc-200 text-zinc-600" : "bg-orange-100 text-orange-600"
                      )}>
                        {m.role === 'user' ? <User className="w-4 h-4" /> : <Bot className="w-4 h-4" />}
                      </div>
                      <div className={cn(
                        "max-w-[80%] p-3 rounded-2xl text-sm",
                        m.role === 'user' ? "bg-orange-600 text-white rounded-tr-none" : "bg-white border border-zinc-200 text-zinc-800 rounded-tl-none shadow-sm"
                      )}>
                        <div className="markdown-body prose prose-sm max-w-none">
                          <ReactMarkdown>{m.text}</ReactMarkdown>
                        </div>
                        {m.sources && m.sources.length > 0 && (
                          <div className="mt-3 pt-3 border-t border-zinc-100">
                            <p className="text-[10px] font-bold text-zinc-400 uppercase mb-2 flex items-center gap-1">
                              <Search className="w-3 h-3" /> Sources
                            </p>
                            <div className="flex flex-wrap gap-2">
                              {m.sources.map((source: any, si: number) => (
                                source.web && (
                                  <a 
                                    key={si}
                                    href={source.web.uri}
                                    target="_blank"
                                    rel="noopener noreferrer"
                                    className="text-[10px] bg-zinc-100 text-zinc-600 px-2 py-1 rounded hover:bg-zinc-200 transition-colors flex items-center gap-1 max-w-full truncate"
                                  >
                                    <span className="truncate">{source.web.title || source.web.uri}</span>
                                  </a>
                                )
                              ))}
                            </div>
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                  {isLoading && (
                    <div className="flex gap-3">
                      <div className="w-8 h-8 rounded-full bg-orange-100 text-orange-600 flex items-center justify-center shrink-0">
                        <Bot className="w-4 h-4" />
                      </div>
                      <div className="bg-white border border-zinc-200 p-3 rounded-2xl rounded-tl-none shadow-sm flex items-center gap-2">
                        <Loader2 className="w-4 h-4 animate-spin text-orange-600" />
                        <span className="text-xs text-zinc-500">Thinking...</span>
                      </div>
                    </div>
                  )}
                </div>

                {/* Input */}
                <div className="p-4 border-t border-zinc-200 bg-white shrink-0">
                  <div className="relative">
                    <input
                      type="text"
                      value={input}
                      onChange={(e) => setInput(e.target.value)}
                      onKeyPress={(e) => e.key === 'Enter' && handleSend()}
                      placeholder="Ask me anything..."
                      className="w-full bg-zinc-50 border border-zinc-200 rounded-xl px-4 py-3 pr-12 text-sm focus:outline-none focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500 transition-all"
                    />
                    <button
                      onClick={handleSend}
                      disabled={!input.trim() || isLoading}
                      className="absolute right-2 top-1/2 -translate-y-1/2 p-2 bg-orange-600 text-white rounded-lg disabled:opacity-50 disabled:hover:bg-orange-600 hover:bg-orange-700 transition-colors"
                    >
                      <Send className="w-4 h-4" />
                    </button>
                  </div>
                  <p className="text-[10px] text-zinc-400 mt-2 text-center">
                    Powered by Gemini 3.1 Flash Lite
                  </p>
                </div>
              </>
            )}
          </motion.div>
        )}
      </AnimatePresence>
    </>
  );
};
