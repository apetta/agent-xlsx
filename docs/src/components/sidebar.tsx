"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { useMobileNav } from "./mobile-nav-context";
import { allDocsPages } from "@/lib/docs-navigation";

export function Sidebar() {
  const pathname = usePathname();
  const { isOpen, setIsOpen } = useMobileNav();

  return (
    <>
      {/* Mobile overlay */}
      {isOpen && (
        <div
          className="lg:hidden fixed inset-0 z-40 bg-background/80 overscroll-contain"
          onClick={() => setIsOpen(false)}
        />
      )}

      {/* Sidebar */}
      <aside
        className={`
          fixed lg:sticky top-14 left-0 z-50 lg:z-auto
          w-56 lg:w-48 h-[calc(100vh-3.5rem)]
          bg-background
          transform transition-transform duration-150 ease-out
          ${isOpen ? "translate-x-0" : "-translate-x-full lg:translate-x-0"}
        `}
      >
        <div className="h-full overflow-y-auto py-5 pl-3 pr-5">
          <nav className="space-y-0.5">
            {allDocsPages.map((item) => {
              const isActive = pathname === item.href;

              return (
                <Link
                  key={item.name}
                  href={item.href}
                  className={`block px-2 py-1.5 text-sm transition-colors ${
                    isActive
                      ? "text-foreground font-medium"
                      : "text-muted-foreground hover:text-foreground"
                  }`}
                >
                  {item.name}
                </Link>
              );
            })}
          </nav>
        </div>
      </aside>
    </>
  );
}
