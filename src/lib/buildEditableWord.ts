import {
  AlignmentType,
  BorderStyle,
  Document,
  ImageRun,
  LevelFormat,
  Packer,
  PageOrientation,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
  convertMillimetersToTwip,
} from "docx";

export interface WordCertificateData {
  projectCode: string;
  projectName: string;
  requester: string;
  requesterContact: string;
  endUser: string;
  deliveryAddress: string;
  completionDate: string;
  serviceEngineer: string[];
  serviceLevel: "urgent" | "non-urgent" | "";
  serviceTypes: string[];
  taskItems: string[];
  hasTimeline: boolean;
  timeline: { bjTime: string; localTime: string; action: string }[];
  result: "completed" | "";
  notes: string;
  confirmedWithRemote: boolean;
  remoteContact: string;
  enriginSig: string;
}

interface WordAssets {
  logo: string;
  signature?: string;
}

const BLUE = "2B5F8B";
const DARK = "1A2E3D";
const MUTED = "6B8099";
const BORDER = "C8D9E8";
const LABEL_FILL = "EDF3F9";
const LIGHT_FILL = "FAFCFF";
const WHITE = "FFFFFF";
const A4_WIDTH = 11906;
const A4_HEIGHT = 16838;
const SIDE_MARGIN = convertMillimetersToTwip(14);
const CONTENT_WIDTH = A4_WIDTH - SIDE_MARGIN * 2;
const CELL_MARGINS = { top: 70, bottom: 70, left: 140, right: 140 };

const gridBorder = {
  style: BorderStyle.SINGLE,
  size: 4,
  color: BORDER,
};

const blueBorder = {
  style: BorderStyle.SINGLE,
  size: 8,
  color: BLUE,
};

const noBorder = {
  style: BorderStyle.NONE,
  size: 0,
  color: WHITE,
};

const gridBorders = {
  top: gridBorder,
  bottom: gridBorder,
  left: gridBorder,
  right: gridBorder,
  insideHorizontal: gridBorder,
  insideVertical: gridBorder,
};

const noBorders = {
  top: noBorder,
  bottom: noBorder,
  left: noBorder,
  right: noBorder,
  insideHorizontal: noBorder,
  insideVertical: noBorder,
};

function dataUrlToBytes(dataUrl: string) {
  const base64 = dataUrl.split(",")[1] || "";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

async function imageSourceToBytes(source: string) {
  if (source.startsWith("data:")) {
    return dataUrlToBytes(source);
  }

  const response = await fetch(source);
  if (!response.ok) {
    throw new Error(`Unable to load Word image asset: ${response.status}`);
  }
  return new Uint8Array(await response.arrayBuffer());
}

function textRun(
  text: string,
  {
    size = 15,
    bold = false,
    color = "1A1A1A",
    italics = false,
  }: { size?: number; bold?: boolean; color?: string; italics?: boolean } = {},
) {
  return new TextRun({
    text,
    font: "Hiragino Sans GB",
    size,
    bold,
    color,
    italics,
  });
}

function paragraph(
  children: TextRun[],
  {
    alignment = AlignmentType.LEFT,
    before = 0,
    after = 0,
    line = 240,
    numbering,
    border,
  }: {
    alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
    before?: number;
    after?: number;
    line?: number;
    numbering?: { reference: string; level: number };
    border?: { top?: typeof gridBorder; bottom?: typeof gridBorder };
  } = {},
) {
  return new Paragraph({
    children,
    alignment,
    spacing: { before, after, line },
    numbering,
    border,
  });
}

function spacer(after = 55) {
  return paragraph([textRun("")], { after, line: 120 });
}

function tableCell({
  children,
  width,
  columnSpan,
  fill,
  borders = gridBorders,
  margins = CELL_MARGINS,
}: {
  children: Paragraph[];
  width: number;
  columnSpan?: number;
  fill?: string;
  borders?: typeof gridBorders;
  margins?: typeof CELL_MARGINS;
}) {
  return new TableCell({
    children,
    width: { size: width, type: WidthType.DXA },
    columnSpan,
    verticalAlign: VerticalAlign.CENTER,
    margins: { ...margins, marginUnitType: WidthType.DXA },
    borders,
    shading: fill
      ? { type: ShadingType.CLEAR, color: "auto", fill }
      : undefined,
  });
}

function sectionTitle(title: string) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    layout: TableLayoutType.FIXED,
    borders: noBorders,
    rows: [
      new TableRow({
        children: [
          tableCell({
            width: CONTENT_WIDTH,
            fill: BLUE,
            borders: noBorders,
            margins: { top: 75, bottom: 75, left: 170, right: 170 },
            children: [
              paragraph([textRun(title, { size: 14, bold: true, color: WHITE })]),
            ],
          }),
        ],
      }),
    ],
  });
}

function labelCell(text: string, width: number) {
  return tableCell({
    width,
    fill: LABEL_FILL,
    children: [
      paragraph([textRun(text, { size: 14, bold: true, color: BLUE })]),
    ],
  });
}

function valueCell(text: string, width: number, columnSpan?: number) {
  return tableCell({
    width,
    columnSpan,
    children: [paragraph([textRun(text || " ", { size: 15 })])],
  });
}

function makeInformationTable(data: WordCertificateData) {
  const columns = [2270, 2889, 2270, CONTENT_WIDTH - 7429];
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: columns,
    layout: TableLayoutType.FIXED,
    borders: gridBorders,
    rows: [
      new TableRow({
        children: [
          labelCell("项目名称 / Project Name", columns[0]),
          valueCell(data.projectName, columns[1] + columns[2] + columns[3], 3),
        ],
      }),
      new TableRow({
        children: [
          labelCell("客户 / Customer", columns[0]),
          valueCell(data.requester, columns[1]),
          labelCell("联系方式 / Contact", columns[2]),
          valueCell(data.requesterContact, columns[3]),
        ],
      }),
      new TableRow({
        children: [
          labelCell("最终用户 / End User", columns[0]),
          valueCell(data.endUser, columns[1]),
          labelCell("完工日期 / Completion Date", columns[2]),
          valueCell(data.completionDate, columns[3]),
        ],
      }),
      new TableRow({
        children: [
          labelCell("交付地址 / Delivery Address", columns[0]),
          valueCell(data.deliveryAddress, columns[1] + columns[2] + columns[3], 3),
        ],
      }),
      new TableRow({
        children: [
          labelCell("服务工程师 / Engineer", columns[0]),
          valueCell(data.serviceEngineer.join("、"), columns[1]),
          labelCell("服务级别 / Level", columns[2]),
          valueCell(
            data.serviceLevel === "urgent"
              ? "紧急 / Urgent"
              : data.serviceLevel === "non-urgent"
                ? "非紧急 / Non-urgent"
                : "",
            columns[3],
          ),
        ],
      }),
      new TableRow({
        children: [
          labelCell("服务类型 / Service Type", columns[0]),
          valueCell(data.serviceTypes.join(" · "), columns[1] + columns[2] + columns[3], 3),
        ],
      }),
    ],
  });
}

function makeTaskTable(data: WordCertificateData) {
  const tasks = data.taskItems.filter(Boolean);
  const children = tasks.length
    ? tasks.map((task) =>
        paragraph([textRun(task, { size: 15 })], {
          after: 25,
          numbering: { reference: "task-list", level: 0 },
        }),
      )
    : [paragraph([textRun("（任务详情 / Task details）", { size: 15, color: "AAAAAA" })])];

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    layout: TableLayoutType.FIXED,
    borders: gridBorders,
    rows: [
      new TableRow({
        children: [
          tableCell({
            width: CONTENT_WIDTH,
            fill: LIGHT_FILL,
            children,
            margins: { top: 130, bottom: 130, left: 220, right: 220 },
          }),
        ],
      }),
    ],
  });
}

function makeTimelineTable(data: WordCertificateData) {
  const rows = data.timeline.filter((row) => row.action || row.bjTime || row.localTime);
  const columns = [2200, 2200, CONTENT_WIDTH - 4400];
  const headerCell = (text: string, width: number) =>
    tableCell({
      width,
      fill: BLUE,
      children: [paragraph([textRun(text, { size: 14, bold: true, color: WHITE })])],
      borders: {
        top: blueBorder,
        bottom: blueBorder,
        left: blueBorder,
        right: blueBorder,
        insideHorizontal: blueBorder,
        insideVertical: blueBorder,
      },
    });

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: columns,
    layout: TableLayoutType.FIXED,
    borders: gridBorders,
    rows: [
      new TableRow({
        children: [
          headerCell("北京时间 / BJ Time", columns[0]),
          headerCell("当地时间 / Local Time", columns[1]),
          headerCell("行动 / Action", columns[2]),
        ],
      }),
      ...(rows.length ? rows : [{ bjTime: "", localTime: "", action: "" }]).map(
        (row, index) =>
          new TableRow({
            children: [
              tableCell({
                width: columns[0],
                fill: index % 2 ? "F4F8FC" : WHITE,
                children: [paragraph([textRun(row.bjTime || "—", { size: 14 })])],
              }),
              tableCell({
                width: columns[1],
                fill: index % 2 ? "F4F8FC" : WHITE,
                children: [paragraph([textRun(row.localTime || "—", { size: 14 })])],
              }),
              tableCell({
                width: columns[2],
                fill: index % 2 ? "F4F8FC" : WHITE,
                children: [paragraph([textRun(row.action || "—", { size: 14 })])],
              }),
            ],
          }),
      ),
    ],
  });
}

function makeResultTable(data: WordCertificateData) {
  const labelWidth = 2270;
  const valueWidth = CONTENT_WIDTH - labelWidth;
  const rows = [
    new TableRow({
      children: [
        labelCell("服务结果 / Result", labelWidth),
        valueCell(data.result === "completed" ? "☑ 已完成 / Completed" : "☐ 已完成 / Completed", valueWidth),
      ],
    }),
    new TableRow({
      children: [
        labelCell("已与客户确认 / Confirmed", labelWidth),
        valueCell(
          `${data.confirmedWithRemote ? "☑" : "☐"} 是 / Yes${
            data.confirmedWithRemote && data.remoteContact ? ` — ${data.remoteContact}` : ""
          }`,
          valueWidth,
        ),
      ],
    }),
  ];

  if (data.notes) {
    rows.push(
      new TableRow({
        children: [
          labelCell("备注 / Notes", labelWidth),
          valueCell(data.notes, valueWidth),
        ],
      }),
    );
  }

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [labelWidth, valueWidth],
    layout: TableLayoutType.FIXED,
    borders: gridBorders,
    rows,
  });
}

function makeSignatureCell(
  label: string,
  sub: string,
  width: number,
  signature?: Uint8Array,
) {
  const children: Paragraph[] = [
    paragraph([textRun(label, { size: 14, bold: true, color: BLUE })], { after: 15 }),
    paragraph([textRun(sub, { size: 13, color: "888888" })], { after: 20 }),
  ];

  if (signature) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 20 },
        children: [
          new ImageRun({
            type: "jpg",
            data: signature,
            transformation: { width: 120, height: 78 },
          }),
        ],
      }),
    );
  } else {
    children.push(spacer(280));
  }

  children.push(
    paragraph([textRun("Authorized Signature", { size: 13, color: "AAAAAA" })], {
      alignment: AlignmentType.CENTER,
      before: 20,
      border: { top: gridBorder },
    }),
  );

  return tableCell({
    width,
    children,
    margins: { top: 140, bottom: 110, left: 170, right: 170 },
  });
}

function makeSignatureTable(signature?: Uint8Array) {
  const gap = 320;
  const boxWidth = Math.floor((CONTENT_WIDTH - gap) / 2);
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [boxWidth, gap, boxWidth],
    layout: TableLayoutType.FIXED,
    borders: noBorders,
    rows: [
      new TableRow({
        children: [
          makeSignatureCell(
            "Enrigin Limited",
            "服务方签章 / Authorized Signature",
            boxWidth,
            signature,
          ),
          tableCell({
            width: gap,
            borders: noBorders,
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            children: [paragraph([textRun("")])],
          }),
          makeSignatureCell(
            "客户 / Customer",
            "客户签章 / Authorized Signature or Chop",
            boxWidth,
          ),
        ],
      }),
    ],
  });
}

function makeHeader(data: WordCertificateData, logo: Uint8Array) {
  const leftWidth = 3100;
  const centerWidth = 4120;
  const rightWidth = CONTENT_WIDTH - leftWidth - centerWidth;
  const bottomOnly = {
    top: noBorder,
    bottom: blueBorder,
    left: noBorder,
    right: noBorder,
    insideHorizontal: noBorder,
    insideVertical: noBorder,
  };

  const left = tableCell({
    width: leftWidth,
    borders: bottomOnly,
    margins: { top: 0, bottom: 130, left: 0, right: 100 },
    children: [
      new Paragraph({
        spacing: { before: 0, after: 20 },
        children: [
          new ImageRun({
            type: "png",
            data: logo,
            transformation: { width: 145, height: 63 },
          }),
        ],
      }),
      paragraph([textRun("英源國際有限公司", { size: 13, color: MUTED })]),
      paragraph([
        textRun(
          "9th Floor, Amtel Building, 148 Des Voeux Road, Central, Hong Kong",
          { size: 12, color: MUTED },
        ),
      ]),
    ],
  });

  const center = tableCell({
    width: centerWidth,
    borders: bottomOnly,
    margins: { top: 80, bottom: 130, left: 100, right: 100 },
    children: [
      paragraph([textRun("完工确认书", { size: 30, bold: true, color: DARK })], {
        alignment: AlignmentType.CENTER,
        after: 10,
      }),
      paragraph(
        [textRun("COMPLETION CERTIFICATE", { size: 20, bold: true, color: BLUE })],
        { alignment: AlignmentType.CENTER },
      ),
    ],
  });

  const right = tableCell({
    width: rightWidth,
    borders: bottomOnly,
    margins: { top: 260, bottom: 130, left: 100, right: 0 },
    children: [
      paragraph(
        [
          textRun("Date / 日期  ", { size: 12, color: MUTED }),
          textRun(data.completionDate || "—", { size: 14, bold: true }),
        ],
        { alignment: AlignmentType.RIGHT, after: 20 },
      ),
      paragraph(
        [
          textRun("Project Code / 项目编号  ", { size: 12, color: MUTED }),
          textRun(data.projectCode || "—", { size: 14, bold: true }),
        ],
        { alignment: AlignmentType.RIGHT },
      ),
    ],
  });

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [leftWidth, centerWidth, rightWidth],
    layout: TableLayoutType.FIXED,
    borders: noBorders,
    rows: [new TableRow({ children: [left, center, right] })],
  });
}

export async function buildEditableWord(
  data: WordCertificateData,
  assets: WordAssets,
) {
  const [logoBytes, signatureBytes] = await Promise.all([
    imageSourceToBytes(assets.logo),
    assets.signature ? imageSourceToBytes(assets.signature) : Promise.resolve(undefined),
  ]);

  const documentChildren: (Paragraph | Table)[] = [
    makeHeader(data, logoBytes),
    spacer(70),
    sectionTitle("SECTION 1 · 项目信息 / Project Information"),
    spacer(20),
    makeInformationTable(data),
    spacer(65),
    sectionTitle("SECTION 2 · 任务详情 / Task Details"),
    spacer(20),
    makeTaskTable(data),
    spacer(65),
  ];

  if (data.hasTimeline) {
    documentChildren.push(
      sectionTitle("SECTION 3 · 服务时间轴 / Service Timeline"),
      spacer(20),
      makeTimelineTable(data),
      spacer(65),
    );
  }

  documentChildren.push(
    sectionTitle(
      `${data.hasTimeline ? "SECTION 4" : "SECTION 3"} · 服务结果 / Service Result`,
    ),
    spacer(20),
    makeResultTable(data),
    spacer(75),
    makeSignatureTable(data.enriginSig ? signatureBytes : undefined),
    spacer(55),
    paragraph(
      [textRun("ENRIGIN LIMITED 英源國際有限公司", { size: 12, color: "AAAAAA" })],
      {
        alignment: AlignmentType.CENTER,
        before: 20,
        border: { top: gridBorder },
      },
    ),
  );

  const word = new Document({
    numbering: {
      config: [
        {
          reference: "task-list",
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: "%1.",
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 420, hanging: 240 },
                  spacing: { after: 25, line: 260 },
                },
              },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              orientation: PageOrientation.PORTRAIT,
              width: A4_WIDTH,
              height: A4_HEIGHT,
            },
            margin: {
              top: convertMillimetersToTwip(12),
              right: SIDE_MARGIN,
              bottom: convertMillimetersToTwip(10),
              left: SIDE_MARGIN,
              header: 0,
              footer: 0,
              gutter: 0,
            },
          },
        },
        children: documentChildren,
      },
    ],
  });

  return Packer.toBlob(word);
}
