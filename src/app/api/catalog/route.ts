import { NextResponse } from "next/server";

export const dynamic = "force-dynamic";

const DEFAULT_SOURCE_URL = "https://zenithpharmaimports.com/api/products";

type SourceProduct = {
  id?: unknown;
  category?: unknown;
  group?: unknown;
  brand?: unknown;
  displayBrand?: unknown;
  name?: unknown;
  presentation?: unknown;
  finalPrice?: unknown;
  status?: unknown;
  isDeleted?: unknown;
};

function optionalText(value: unknown) {
  return typeof value === "string" ? value.trim() : "";
}

function normalizeProduct(value: SourceProduct) {
  const id = optionalText(value.id);
  const category = optionalText(value.category);
  const name = optionalText(value.name);
  const presentation = optionalText(value.presentation);
  const displayBrand = optionalText(value.displayBrand);
  const nameBrand = name.includes(" - ") ? name.split(" - ").at(-1)?.trim() || "" : "";
  const brand = optionalText(value.brand) || displayBrand || nameBrand || "Sem marca";

  if (!id || !category || !name) return null;

  const parsedPrice = value.finalPrice === null || value.finalPrice === undefined || value.finalPrice === ""
    ? null
    : typeof value.finalPrice === "number"
      ? value.finalPrice
      : Number(value.finalPrice);

  return {
    id,
    category,
    group: optionalText(value.group) || category,
    brand,
    displayBrand,
    name,
    presentation,
    finalPrice: parsedPrice !== null && Number.isFinite(parsedPrice) ? parsedPrice : null,
    status: value.status === "out_of_stock" || value.status === "inactive"
      ? value.status
      : "active",
    isDeleted: value.isDeleted === true,
  };
}

export async function GET() {
  const sourceUrl = process.env.ZENITH_CATALOG_SOURCE_URL || DEFAULT_SOURCE_URL;

  try {
    const response = await fetch(sourceUrl, {
      cache: "no-store",
      headers: {
        accept: "application/json",
        "user-agent": "Zenith-Minimal-Catalog/1.0",
      },
    });

    if (!response.ok) {
      throw new Error(`Fonte respondeu com HTTP ${response.status}.`);
    }

    const payload: unknown = await response.json();
    if (!Array.isArray(payload)) {
      throw new Error("A fonte não retornou uma lista de produtos.");
    }

    const products = payload
      .map((item) => normalizeProduct(item as SourceProduct))
      .filter((item) => item !== null && !item.isDeleted && item.status !== "inactive");

    if (products.length === 0) {
      throw new Error("A fonte retornou um catálogo vazio.");
    }

    return NextResponse.json(
      {
        source: "zenithpharmaimports.com",
        syncedAt: new Date().toISOString(),
        products,
      },
      {
        headers: {
          "Cache-Control": "no-store, max-age=0, must-revalidate",
        },
      },
    );
  } catch (error) {
    console.error("Falha ao sincronizar o catálogo Zenith:", error);
    return NextResponse.json(
      { error: "Não foi possível consultar o catálogo oficial agora." },
      {
        status: 502,
        headers: {
          "Cache-Control": "no-store, max-age=0, must-revalidate",
        },
      },
    );
  }
}
