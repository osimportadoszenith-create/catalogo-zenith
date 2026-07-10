import { NextRequest, NextResponse } from "next/server";

export function middleware(request: NextRequest) {
  const expected = process.env.CATALOG_ACCESS_TOKEN;
  const token = request.nextUrl.searchParams.get("acesso");

  if (!expected || token !== expected) {
    return new NextResponse(null, { status: 404 });
  }

  return NextResponse.next();
}

export const config = {
  matcher: ["/", "/catalogo-zenith.html"],
};
