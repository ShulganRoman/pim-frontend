import * as XLSX from "xlsx";

export const PRODUCT_SHEETS = [
  "TCT_Router_Bit",
  "Insert_Tool",
  "Countersink",
];

export const SHEETS = {
  README: "README",
  CONFIG: "Import_Config",
  GROUPS: "Attribute_Groups",
  ATTRIBUTES: "Attributes",
  TYPES: "Types",
  TYPE_GROUP_BINDINGS: "Type_Group_Bindings",
  ITEM_PARENTS: "Item_Parents",
};

export const CONFIG_HEADERS = ["key", "value"];

export const GROUP_HEADERS = [
  "identifier",
  "name_en",
  "order",
  "visible",
  "options_json",
];

export const ATTRIBUTE_HEADERS = [
  "identifier",
  "name_en",
  "type_code",
  "groups_csv",
  "order",
  "language_dependent",
  "rich_text",
  "multi_line",
  "pattern",
  "lov_identifier",
  "options_json",
  "valid_types_csv",
  "visible_types_csv",
];

export const TYPE_HEADERS = [
  "identifier",
  "name_en",
  "parent_identifier",
  "icon",
  "icon_color",
  "file",
];

export const TYPE_GROUP_BINDING_HEADERS = [
  "group_identifier",
  "type_identifier",
  "valid",
  "visible",
];

export const ITEM_BASE_HEADERS = [
  "identifier",
  "name_en",
  "type_identifier",
  "parent_identifier",
  "values_json",
  "channels_json",
];

export const ITEM_PARENT_HEADERS = [...ITEM_BASE_HEADERS];

export const PRODUCT_HEADERS = [
  ...ITEM_BASE_HEADERS,
  "attr:cutting_diameter",
];

const TYPE_CODE_HELP = [
  "1=TEXT",
  "2=BOOLEAN",
  "3=INTEGER",
  "4=FLOAT",
  "5=DATE",
  "6=TIME",
  "7=ENUM",
  "8=URL",
].join(", ");

function addSheet(workbook, name, rows) {
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, name);
}

export function createImportTemplateWorkbook() {
  const workbook = XLSX.utils.book_new();

  addSheet(workbook, SHEETS.README, [
    ["PIM Excel Import Template"],
    [""],
    ["Required product sheets:"],
    [PRODUCT_SHEETS.join(", ")],
    [""],
    ["Fill metadata sheets first:"],
    [SHEETS.GROUPS],
    [SHEETS.ATTRIBUTES],
    [SHEETS.TYPES],
    [SHEETS.TYPE_GROUP_BINDINGS],
    [SHEETS.ITEM_PARENTS],
    [""],
    ["Rules:"],
    ["1) All identifiers must be unique and stable."],
    [
      "2) Create or reference parent items for child types. Child-type products require parent_identifier.",
    ],
    [
      "3) Product values can be set through values_json or through dynamic columns attr:<attribute_identifier>.",
    ],
    ["4) values_json and channels_json must contain valid JSON objects."],
    ["5) Attribute type_code values:"],
    [TYPE_CODE_HELP],
    ["6) Booleans accept: true/false/1/0/yes/no."],
    [
      "7) Type_Group_Bindings sheet is used to propagate type visibility/validity to attributes by group.",
    ],
    [
      "8) Item_Parents rows are imported before product sheets so product rows can reference them.",
    ],
  ]);

  addSheet(workbook, SHEETS.CONFIG, [
    CONFIG_HEADERS,
    ["mode", "CREATE_UPDATE"],
    ["errors", "PROCESS_WARN"],
    ["default_language", "en"],
  ]);

  addSheet(workbook, SHEETS.GROUPS, [
    GROUP_HEADERS,
    ["cutting_geometry", "Cutting Geometry", 10, "TRUE", "{}"],
    ["commercial", "Commercial Data", 20, "TRUE", "{}"],
  ]);

  addSheet(workbook, SHEETS.ATTRIBUTES, [
    ATTRIBUTE_HEADERS,
    [
      "cutting_diameter",
      "Cutting Diameter",
      4,
      "cutting_geometry",
      10,
      "FALSE",
      "FALSE",
      "FALSE",
      "",
      "",
      "{}",
      "",
      "",
    ],
    [
      "material",
      "Material",
      1,
      "commercial",
      20,
      "FALSE",
      "FALSE",
      "FALSE",
      "",
      "",
      "{}",
      "",
      "",
    ],
    [
      "is_coated",
      "Is Coated",
      2,
      "commercial",
      30,
      "FALSE",
      "FALSE",
      "FALSE",
      "",
      "",
      "{}",
      "",
      "",
    ],
  ]);

  addSheet(workbook, SHEETS.TYPES, [
    TYPE_HEADERS,
    ["product_type", "Product Type", "", "shape-outline", "blue", "FALSE"],
    ["tct_router_bit", "TCT Router Bit", "product_type", "saw-blade", "indigo", "FALSE"],
    ["insert_tool", "Insert Tool", "product_type", "tools", "teal", "FALSE"],
    ["countersink", "Countersink", "product_type", "drill", "orange", "FALSE"],
  ]);

  addSheet(workbook, SHEETS.TYPE_GROUP_BINDINGS, [
    TYPE_GROUP_BINDING_HEADERS,
    ["cutting_geometry", "tct_router_bit", "TRUE", "TRUE"],
    ["cutting_geometry", "insert_tool", "TRUE", "TRUE"],
    ["cutting_geometry", "countersink", "TRUE", "TRUE"],
    ["commercial", "tct_router_bit", "TRUE", "TRUE"],
    ["commercial", "insert_tool", "TRUE", "TRUE"],
    ["commercial", "countersink", "TRUE", "TRUE"],
  ]);

  addSheet(workbook, SHEETS.ITEM_PARENTS, [
    ITEM_PARENT_HEADERS,
    ["catalog_root_001", "Catalog Root 001", "product_type", "", "{}", "{}"],
  ]);

  addSheet(workbook, PRODUCT_SHEETS[0], [
    PRODUCT_HEADERS,
    [
      "router_bit_001",
      "Router Bit 001",
      "tct_router_bit",
      "catalog_root_001",
      '{"material":"Carbide","is_coated":true}',
      "{}",
      12.7,
    ],
  ]);

  addSheet(workbook, PRODUCT_SHEETS[1], [
    PRODUCT_HEADERS,
    [
      "insert_tool_001",
      "Insert Tool 001",
      "insert_tool",
      "catalog_root_001",
      '{"material":"Steel","is_coated":false}',
      "{}",
      8.5,
    ],
  ]);

  addSheet(workbook, PRODUCT_SHEETS[2], [
    PRODUCT_HEADERS,
    [
      "countersink_001",
      "Countersink 001",
      "countersink",
      "catalog_root_001",
      '{"material":"HSS","is_coated":true}',
      "{}",
      6.2,
    ],
  ]);

  return workbook;
}

export function downloadImportTemplate() {
  const workbook = createImportTemplateWorkbook();
  XLSX.writeFile(workbook, "PIM_Import_Template.xlsx");
}
