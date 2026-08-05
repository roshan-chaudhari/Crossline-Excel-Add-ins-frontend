// Mock global browser and Office environments before requiring taskpane.js
global.Office = {
  onReady: jest.fn()
};
global.document = {
  getElementById: jest.fn().mockReturnValue({
    addEventListener: jest.fn(),
    style: {},
    innerText: "",
    onclick: null
  }),
  addEventListener: jest.fn()
};
global.window = {};
const { buildFieldMap, mapToCostingUpdateRequest } = require("./taskpane.js");
describe("buildFieldMap", () => {
  it("should generate Knit Costing fields correctly when costingModel is Knit / Sweater Costing V1", () => {
    const data = [
      {
        costingModel: "Knit / Sweater Costing V1",
        styleNumber: "KS260044",
        fabricModel: [
          {
            fabricNumber: "KF-4045",
            costPrice: 5.5,
            quantity: 1.89
          }
        ],
        trimModel: [
          {
            trimMaterialCategory: "Sewing / Labeling Accessories",
            trimNumber: "T00006",
            costPrice: 0.40,
            quantity: 1.00
          }
        ]
      }
    ];
    const fields = buildFieldMap(data);
    // Verify SupplierId has correct label
    const supplierIdField = fields.find(f => f.key === "supplierId");
    expect(supplierIdField).toBeDefined();
    expect(supplierIdField.label).toBe("SupplierId");

    // Verify BOM Yarn Cost exists instead of BOM Fabric Cost
    const hasBomYarnCost = fields.some(f => f.label === "BOM Yarn Cost");
    const hasBomFabricCost = fields.some(f => f.label === "BOM Fabric Cost");
    expect(hasBomYarnCost).toBe(true);
    expect(hasBomFabricCost).toBe(false);
    // Verify Yarn Information is in section 2
    const yarnInfoFields = fields.filter(f => f.section === "2 - Yarn Information");
    expect(yarnInfoFields.length).toBeGreaterThan(0);
    // Verify standard category default slots exist even if no data is present
    // For example, "Wash" category should have a blank slot at slot 0
    const washFields = fields.filter(f => f._trimGroupCategory === "Wash" && f._trimGroupSlot === 0);
    expect(washFields.length).toBe(5); // Number, Name, Price, Consumption, UOM
  });
  it("should generate Woven Costing fields correctly when costingModel is Woven Costing V1", () => {
    const data = [
      {
        costingModel: "Woven Costing V1",
        styleNumber: "WS260046",
        fabricModel: [
          {
            fabricNumber: "FW0V00002",
            costPrice: 2.25,
            quantity: 19.75
          }
        ]
      }
    ];
    const fields = buildFieldMap(data);
    // Verify BOM Fabric Cost exists instead of BOM Yarn Cost
    const hasBomYarnCost = fields.some(f => f.label === "BOM Yarn Cost");
    const hasBomFabricCost = fields.some(f => f.label === "BOM Fabric Cost");
    expect(hasBomYarnCost).toBe(false);
    expect(hasBomFabricCost).toBe(true);
    // Verify Shell Fabric Information is in section 2
    const shellFabricFields = fields.filter(f => f.section === "2 - Shell Fabric Information");
    expect(shellFabricFields.length).toBeGreaterThan(0);
  });

  it("should group fabricModel item with materialCategory='Print & Embroidery' under Print & Embroidery Information in Woven/Knit models", () => {
    const data = [
      {
        costingModel: "Woven Costing V1",
        styleNumber: "WS260047",
        fabricModel: [
          {
            fabricNumber: "FWOV00004",
            costPrice: 2.25,
            quantity: 19.75
          },
          {
            fabricNumber: "PPRN00001",
            fabricName: "Print",
            costPrice: 4,
            quantity: 1,
            uom: "Dozen",
            materialCategory: "Print & Embroidery"
          }
        ]
      }
    ];

    // Preprocess like in taskpane.js fetch callback
    data.forEach((costing) => {
      if (costing.fabricModel) {
        costing.printEmbroideryFabrics = costing.fabricModel.filter(
          (f) => f && (f.materialCategory || "").trim() === "Print & Embroidery"
        );
        costing.fabricModel = costing.fabricModel.filter(
          (f) => f && (f.materialCategory || "").trim() !== "Print & Embroidery"
        );
      } else {
        costing.printEmbroideryFabrics = [];
      }
    });

    const fields = buildFieldMap(data);

    // Verify fabricModel only contains the non-Print-&-Embroidery fabric
    expect(data[0].fabricModel.length).toBe(1);
    expect(data[0].fabricModel[0].fabricNumber).toBe("FWOV00004");
    expect(data[0].printEmbroideryFabrics.length).toBe(1);
    expect(data[0].printEmbroideryFabrics[0].fabricNumber).toBe("PPRN00001");
    // Check that Print & Embroidery category contains the fabric item
    const peFields = fields.filter(f => f._trimGroupCategory === "Print & Embroidery");
    expect(peFields.length).toBeGreaterThan(0);
    // Verify trimIndicesPerCol contains the fabric reference
    const firstPeField = peFields[0];
    const colIndices = firstPeField._trimIndicesPerCol[0];
    expect(colIndices).toContainEqual({ source: "fabric", index: 0 });
  });
});

describe("mapToCostingUpdateRequest", () => {
  it("should prepend the main fabric/yarn info to the FabricModel list when mainFabricFabricName is present", () => {
    const costingObj = {
      styleNumber: "WS260041",
      season: "S2026",
      supplierId: 67,
      bomVersion: "BLEACH",
      mainFabricFabricName: "MAIN_FAB_001",
      mainFabricCostPrice: 5.5,
      mainFabricQuantity: 1.25,
      fabricModel: [
        {
          fabricNumber: "SUB_FAB_002",
          costPrice: 2.5,
          quantity: 0.75
        }
      ],
      trimModel: [
        {
          trimNumber: "TRIM_001",
          costPrice: 0.15,
          quantity: 2
        }
      ]
    };

    const payload = mapToCostingUpdateRequest("Costing 2", 3, costingObj);

    expect(payload["Costing Name"]).toBe("Costing 2");
    expect(payload["ColIndex"]).toBe("3");
    expect(payload["StyleNumber"]).toBe("WS260041");
    expect(payload["FabricModel"]).toHaveLength(2);

    // Verify main fabric is the first element
    expect(payload["FabricModel"][0]).toEqual({
      "Fabric Number": "MAIN_FAB_001",
      "Cost Price": "5.5",
      "Quantity": "1.25"
    });

    // Verify other fabric is the second element
    expect(payload["FabricModel"][1]).toEqual({
      "Fabric Number": "SUB_FAB_002",
      "Cost Price": "2.5",
      "Quantity": "0.75"
    });

    // Verify trim model
    expect(payload["TrimModel"]).toHaveLength(1);
    expect(payload["TrimModel"][0]).toEqual({
      "Trim Number": "TRIM_001",
      "Cost Price": "0.15",
      "Quantity": "2"
    });
  });

  it("should not prepend main fabric/yarn if mainFabricFabricName is not present", () => {
    const costingObj = {
      styleNumber: "WS260041",
      season: "S2026",
      supplierId: 67,
      bomVersion: "BLEACH",
      mainFabricFabricName: "", // empty
      mainFabricCostPrice: 0,
      mainFabricQuantity: 0,
      fabricModel: [
        {
          fabricNumber: "SUB_FAB_002",
          costPrice: 2.5,
          quantity: 0.75
        }
      ]
    };

    const payload = mapToCostingUpdateRequest("Costing 2", 3, costingObj);

    expect(payload["FabricModel"]).toHaveLength(1);
    expect(payload["FabricModel"][0]).toEqual({
      "Fabric Number": "SUB_FAB_002",
      "Cost Price": "2.5",
      "Quantity": "0.75"
    });
  });
});

