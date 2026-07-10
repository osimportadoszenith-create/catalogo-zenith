import { NextRequest, NextResponse } from "next/server";

// O catálogo só é acessível via /py (ver next.config.js). O nome de
// arquivo real fica bloqueado para não vazar como caminho alternativo.
export function middleware(_request: NextRequest) {
  return new NextResponse(null, { status: 404 });
}

export const config = {
  matcher: ["/catalogo-zenith.html"],
};
