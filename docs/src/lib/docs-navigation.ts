export type NavItem = {
  name: string;
  href: string;
};

export const allDocsPages: NavItem[] = [
  { name: "Introduction", href: "/" },
  { name: "Installation", href: "/installation" },
  { name: "Quick Start", href: "/quick-start" },
  { name: "Commands", href: "/commands" },
  { name: "Backends", href: "/backends" },
  { name: "Formats", href: "/formats" },
  { name: "Changelog", href: "/changelog" },
];
