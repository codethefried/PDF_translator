import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { Languages } from "lucide-react";
import { HomePage } from "@/pages/home";

const queryClient = new QueryClient({ defaultOptions: { queries: { refetchOnWindowFocus: false } } });

export function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <div className="flex min-h-screen flex-col">
        <a href="#result" className="sr-only focus:not-sr-only focus:absolute focus:left-3 focus:top-3 focus:z-50 focus:rounded-md focus:bg-background focus:px-3 focus:py-2 focus:ring-2 focus:ring-ring">
          번역 결과로 건너뛰기
        </a>
        <header className="sticky top-0 z-30 border-b bg-primary text-primary-foreground">
          <div className="mx-auto flex h-12 max-w-[1500px] items-center gap-2 px-4">
            <Languages className="size-5" aria-hidden="true" />
            <span className="text-base font-semibold tracking-tight">PDF 문서 번역기</span>
          </div>
        </header>
        <main className="flex flex-1 flex-col">
          <HomePage />
        </main>
      </div>
    </QueryClientProvider>
  );
}
