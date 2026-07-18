import React, { Component, ErrorInfo, ReactNode } from "react";
import { AlertOctagon, RotateCcw, Trash2, ChevronDown, ChevronUp } from "lucide-react";

interface Props {
  children: ReactNode;
  fallbackTitle?: string;
  onReset?: () => void;
}

interface State {
  hasError: boolean;
  error: Error | null;
  errorInfo: ErrorInfo | null;
  showDetails: boolean;
}

export class ErrorBoundary extends React.Component<Props, State> {
  public state: State = {
    hasError: false,
    error: null,
    errorInfo: null,
    showDetails: false,
  };

  public static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error, errorInfo: null, showDetails: false };
  }

  public componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error("ErrorBoundary caught an error:", error, errorInfo);
    this.setState({ errorInfo });
  }

  private handleRetry = () => {
    if (this.props.onReset) {
      this.props.onReset();
    }
    this.setState({
      hasError: false,
      error: null,
      errorInfo: null,
      showDetails: false,
    });
    // Force soft refresh of page state without full browser reload
    window.location.hash = "";
  };

  private handleFullReset = () => {
    localStorage.clear();
    sessionStorage.clear();
    window.location.reload();
  };

  public render() {
    if (this.state.hasError) {
      return (
        <div className="p-6 max-w-2xl mx-auto my-8 bg-white border-4 border-[#141414] shadow-[8px_8px_0px_#141414] animate-in fade-in zoom-in-95 duration-150">
          <div className="bg-red-500 border-b-4 border-[#141414] p-4 -mx-6 -mt-6 flex items-center gap-3 text-white">
            <AlertOctagon size={28} className="stroke-[2.5]" />
            <div>
              <h2 className="font-black text-sm uppercase tracking-tight">HỆ THỐNG PHÁT HIỆN LỖI RUNTIME</h2>
              <p className="text-[10px] opacity-90 uppercase font-bold">Ứng dụng đã được bảo vệ tránh hiện tượng trắng trang</p>
            </div>
          </div>

          <div className="mt-6 space-y-4">
            <div className="bg-amber-50 border-2 border-amber-300 p-4 text-xs space-y-2">
              <p className="font-black text-amber-900 uppercase">⚠️ {this.props.fallbackTitle || "ĐÃ XẢY RA LỖI KHI XỬ LÝ DỮ LIỆU EXCEL"}</p>
              <p className="text-slate-700 font-bold leading-relaxed">
                Hệ thống gặp sự cố không mong muốn trong quá trình xử lý cấu trúc tệp dữ liệu của bạn. Điều này thường do tệp Excel có cấu trúc tiêu đề phức tạp hoặc không hợp lệ.
              </p>
            </div>

            {/* Action Buttons */}
            <div className="flex flex-wrap gap-3 pt-2">
              <button
                onClick={this.handleRetry}
                className="flex items-center gap-2 bg-[#00ff00] hover:bg-[#05e005] text-black text-xs font-black uppercase tracking-wider px-4 py-2.5 border-2 border-[#141414] shadow-[3px_3px_0px_#141414] active:translate-y-[2px] active:shadow-[1px_1px_0px_#141414] transition cursor-pointer"
              >
                <RotateCcw size={14} className="stroke-[2.5]" />
                Thử lại ngay
              </button>
              
              <button
                onClick={this.handleFullReset}
                className="flex items-center gap-2 bg-red-100 hover:bg-red-200 text-red-700 text-xs font-black uppercase tracking-wider px-4 py-2.5 border-2 border-red-300 shadow-[3px_3px_0px_#ef4444] active:translate-y-[2px] active:shadow-[1px_1px_0px_#ef4444] transition cursor-pointer"
              >
                <Trash2 size={14} />
                Xóa sạch bộ nhớ & Tải lại
              </button>
            </div>

            {/* Error Detail Accordion */}
            <div className="border-2 border-slate-200 rounded">
              <button
                onClick={() => this.setState(prev => ({ showDetails: !prev.showDetails }))}
                className="w-full flex justify-between items-center bg-slate-100 p-2.5 text-left text-[11px] font-bold uppercase text-slate-600 hover:bg-slate-200 transition"
              >
                <span>Xem chi tiết lỗi kĩ thuật</span>
                {this.state.showDetails ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
              </button>
              
              {this.state.showDetails && (
                <div className="p-3 bg-slate-900 text-slate-200 font-mono text-[10px] overflow-auto max-h-[250px] space-y-2 select-text">
                  <p className="text-red-400 font-bold border-b border-slate-700 pb-1">
                    Error: {this.state.error?.message || "Unknown error"}
                  </p>
                  {this.state.error?.stack && (
                    <pre className="whitespace-pre-wrap leading-normal opacity-85">
                      {this.state.error.stack}
                    </pre>
                  )}
                  {this.state.errorInfo?.componentStack && (
                    <div className="border-t border-slate-800 pt-2 opacity-75">
                      <p className="font-bold uppercase text-slate-400">Component Stack:</p>
                      <pre className="whitespace-pre-wrap leading-normal">
                        {this.state.errorInfo.componentStack}
                      </pre>
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}
