let handlerAttached = false; // global flag
let isRunning = false;
let sheetChangeDebounce = null;

Office.onReady(() => {
  Excel.run(async (context) => {
    if (!handlerAttached) {
      context.workbook.worksheets.onChanged.add(onSheetChanged);
      handlerAttached = true;
    }
    await context.sync();
  });

  if (!StyleResponseData && typeof Office !== "undefined" && Office.context?.document?.settings) {
    try {
      const savedData = Office.context.document.settings.get("StyleResponseData");
      if (savedData) {
        StyleResponseData = typeof savedData === "string" ? JSON.parse(savedData) : savedData;
        console.log("Loaded StyleResponseData from Office document settings on startup:", StyleResponseData);
      }
    } catch (e) {
      console.error("Failed to load StyleResponseData from Office settings on startup:", e);
    }
  }
  toggleButtonVisibility();
});

let ServerURL = process.env.CLIENT_SERVER_URL;
let apiURL = process.env.CLIENT_API_URL;

if (!ServerURL || !apiURL) {
  if (typeof window !== "undefined" && window.location.port === "3000") {
    ServerURL = "https://localhost:3000";
    apiURL = "https://localhost:4000";
  } else {
    ServerURL = typeof window !== "undefined" ? window.location.origin : "https://localhost:3000";
    apiURL = typeof window !== "undefined" ? window.location.origin : "https://localhost:4000";
  }
}

// let Server = "";
// Server = "LocalHost"; //Localhost
// // Server = "Server"; //Server
// let ServerURL = "",
//   apiURL = "";
// if (Server !== "LocalHost") {
//   ServerURL = "https://crosslinecosting-dhemhcdyhcfje5e3.centralindia-01.azurewebsites.net";
// } else {
//   ServerURL = "https://localhost:3000";
//   apiURL = "https://localhost:4000";
// }
let Seasons = [];
let Divisions = [];
let Brands = [];
let Gender = [];
let Category = [];
let tokenExpiryDate = null;
let StyleResponseData = null;
let lastDebugLog = "";
try {
  if (typeof localStorage !== "undefined") {
    const cachedData = localStorage.getItem("StyleResponseData");
    if (cachedData) {
      StyleResponseData = JSON.parse(cachedData);
    }
  }
} catch (e) {
  console.error("Failed to load StyleResponseData from localStorage:", e);
}
document.getElementById("openDialog").addEventListener("click", () => {
  Office.context.ui.displayDialogAsync(
    ServerURL + "/commands.html",
    { height: 70, width: 50, displayInIframe: true },
    function (asyncResult) {
      const dialog = asyncResult.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
        const code = arg.message;
        dialog.close();
        document.getElementById("fileInfo").innerText = code;
        const data = code.split("|");
        const dci = data[0].split("==");
        const dcs = data[1].split("==");
        const dru = data[2].split("==");
        const auc = data[3].split("==");
        const dpu = data[4].split("==");
        const dot = data[5].split("==");
        const dapi = data[6].split("==");
        const dti = data[7].split("==");
        const payload = {
          ci: dci[1],
          cs: dcs[1],
          ru: dru[1],
          code: auc[1],
          pu: dpu[1],
          ot: dot[1],
          api: dapi[1],
          ti: dti[1],
        };

        try {
          const address = await fetch(apiURL + "/api/Auth/token", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
          });

          const token = await address.json();
          var messageObject = {
            messageType: "token",
            access_token: token.access_token,
            time: token.expires_in,
          };
          var jsonMessage = JSON.stringify(messageObject);
          console.log("Token: ", jsonMessage);

          document.getElementById("access_token").innerText = token.access_token;
          if (token.access_token) {
            document.getElementById("token_expiresin").style.display = "inline-block";
            document.getElementById("StyleData").style.display = "inline-block";
            document.getElementById("UploadData").style.display = "none";
            document.getElementById("filtersDiv").style.display = "inline-block";
            const expiryTime = getExpiryTime(token.expires_in);
            tokenExpiryDate = expiryTime;
            document.getElementById("token_expiresin").innerText =
              "Token will expire at: " + expiryTime.toLocaleString();
            const accessToken = document.getElementById("access_token").innerText;
            const seasonUrl = "/api/FashionPlm/seasons";
            const divisionUrl = "/api/FashionPlm/divisions";
            const brandUrl = "/api/FashionPlm/brands";
            const categoryUrl = "/api/FashionPlm/categories";
            const genderUrl = "/api/FashionPlm/genders";
            // const styleListUrl = "/api/FashionPlm/StyleLists";
            const genericUrl = "/api/FashionPlm/Generic";

            fetch(apiURL + seasonUrl, {
              headers: {
                Authorization: `Bearer ${accessToken}`,
                Accept: "application/json",
              },
            })
              .then((res) => res.json())
              .then((data) => {
                const seasons = data.value || [];
                // Sorted Season Code (Id) in ASCENDING order
                const sortedSeasons = seasons.sort((a, b) => {
                  return a.Name.localeCompare(b.Name);
                });
                const seasData = sortedSeasons.map((s) => ({
                  id: s.Id,
                  name: s.Name,
                }));
                Seasons = seasData;
                renderSeasonDropdownOptions(Seasons);
                const dropdown = document.getElementById("seasonSelect");
              });
            fetch(apiURL + divisionUrl, {
              headers: {
                Authorization: `Bearer ${accessToken}`,
                Accept: "application/json",
              },
            })
              .then((res) => res.json())
              .then((data) => {
                const divisions = data.value || [];
                const sortedDivision = divisions.sort((a, b) => {
                  return a.Name.localeCompare(b.Name);
                });
                const divData = sortedDivision.map((s) => ({
                  id: s.Id,
                  name: s.Name,
                }));
                Divisions = divData;
                renderdivisionDropdownOptions(Divisions);
                const dropdown = document.getElementById("divisionSelect");
              });
            fetch(apiURL + brandUrl, {
              headers: {
                Authorization: `Bearer ${accessToken}`,
                Accept: "application/json",
              },
            })
              .then((res) => res.json())
              .then((data) => {
                const brands = data.value || [];
                const sortedBrand = brands.sort((a, b) => {
                  return a.Name.localeCompare(b.Name);
                });
                const brandData = sortedBrand.map((s) => ({
                  id: s.Id,
                  name: s.Name,
                }));
                Brands = brandData;
                renderbrandDropdownOptions(Brands);
                const dropdown = document.getElementById("brandSelect");
              });
            fetch(apiURL + genericUrl, {
              headers: {
                Authorization: `Bearer ${accessToken}`,
                Accept: "application/json",
              },
            })
              .then((res) => res.json())
              .then((data) => {
                //         let data = await response.json();
                const GenericData = data.genericData?.[0];
                if (!GenericData) {
                  throw new Error("No genericData found");
                }
                const headers = Object.keys(GenericData); // ["colorSubType", "part", "udf4"]
                const maxLength = Math.max(...headers.map((h) => GenericData[h].length));
                const rows = [];
                for (let i = 0; i < maxLength; i++) {
                  const row = headers.map((h) => GenericData[h][i] || ""); // fill with empty if missing
                  rows.push(row);
                }
                // Send to Excel
                // writeGenericToSheet(headers, rows);
              });
            fetch(apiURL + categoryUrl, {
              headers: {
                Authorization: `Bearer ${accessToken}`,
                Accept: "application/json",
              },
            })
              .then((res) => res.json())
              .then((data) => {
                const category = data.value || [];
                // Sorted Season Code (Id) in ASCENDING order
                const sortedCategory = category.sort((a, b) => {
                  return a.Name.localeCompare(b.Name);
                });
                const cateData = sortedCategory.map((s) => ({
                  id: s.Id,
                  name: s.Name,
                }));
                Category = cateData;
                renderCategoryDropdownOptions(Category);
                const dropdown = document.getElementById("categorySelect");
              });
            fetch(apiURL + genderUrl, {
              headers: {
                Authorization: `Bearer ${accessToken}`,
                Accept: "application/json",
              },
            })
              .then((res) => res.json())
              .then((data) => {
                const genders = data.value || [];
                // Sorted Season Code (Id) in ASCENDING order
                const sortedGender = genders.sort((a, b) => {
                  return a.Name.localeCompare(b.Name);
                });
                const gendData = sortedGender.map((s) => ({
                  id: s.Id,
                  name: s.Name,
                }));
                Gender = gendData;
                renderGenderDropdownOptions(Gender);
                const dropdown = document.getElementById("genderSelect");
              });
          } else {
            document.getElementById("StyleData").style.display = "none";
            document.getElementById("filtersDiv").style.display = "none";
          }
        } catch (err) {
          console.error("Fetch to /api/token failed:", err);
        }
      });
    }
  );
});
async function FetchAndShowStylesInDropDown() {
  const selectedSeasonId = document.getElementById("selectedSeasons").innerText;
  const selectedDivisionId = document.getElementById("selectedDivisions").innerText;
  const selectedBrandId = document.getElementById("selectedBrands").innerText;
  const selectedGenderId = document.getElementById("selectedGender").innerText;
  const selectedCategoryId = document.getElementById("selectedCategories").innerText;
  // Reset dropdown if any filter is missing
  if (!selectedSeasonId || selectedSeasonId === "None" || !selectedDivisionId || selectedDivisionId === "None" || !selectedBrandId || selectedBrandId === "None" || !selectedGenderId || selectedGenderId === "None" || !selectedCategoryId || selectedCategoryId === "None") {
    const container = document.getElementById("styleOptions");
    if (container) container.innerHTML = '';
    const btn = document.getElementById("styleDropdownBtn");
    if (btn) {
      btn.title = "--Select Style--";
    }
    const valSpan = document.getElementById("styleDropdownValue");
    if (valSpan) {
      valSpan.textContent = "--Select Style--";
    }
    const selectAllCheckbox = document.getElementById("styleSelectAll");
    if (selectAllCheckbox) selectAllCheckbox.checked = false;
    const display = document.getElementById("selectedStyles");
    if (display) {
      display.textContent = "None";
      display.hidden = true;
    }
    return;
  }
  try {
    const accessToken = document.getElementById("access_token").innerText;
    const styleurl = apiURL + `/api/StyleLists/GetStyles?seasonId=${selectedSeasonId}&brandId=${selectedBrandId}&divisionId=${selectedDivisionId}&genderId=${selectedGenderId}&categoryId=${selectedCategoryId}`;
    const res = await fetch(styleurl, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    });
    console.log("Fetching styles with URL:", styleurl);
    const rawData = await res.json();
    console.log("Raw style data:", rawData);
    const Styles = rawData.value || [];
    console.log("Fetched styles:", Styles);

    Styles.sort((a, b) => {
      const codeA = (a.StyleCode ?? "").toLowerCase();
      const codeB = (b.StyleCode ?? "").toLowerCase();
      return codeA.localeCompare(codeB);
    });
    renderStyleDropdownOptions(Styles);
  } catch (error) {
    console.error("Error fetching styles:", error);
  }
}
// });
function renderSeasonDropdownOptions(SeasonData) {
  const container = document.getElementById("seasonSelect");
  container.innerHTML = '<option value="">-- Select Season --</option>'; // Clear existing
  SeasonData.forEach((season) => {
    const option = document.createElement("option");
    option.value = season.id;
    option.textContent = season.name;
    container.appendChild(option);
  });
}
function renderdivisionDropdownOptions(DivisionData) {
  const container = document.getElementById("divisionSelect");
  container.innerHTML = '<option value="">-- Select Department --</option>'; // Clear existing
  DivisionData.forEach((division) => {
    const option = document.createElement("option");
    option.value = division.id;
    option.textContent = division.name;
    container.appendChild(option);
  });
}
function renderbrandDropdownOptions(BrandData) {
  const container = document.getElementById("brandSelect");
  container.innerHTML = '<option value="">-- Select Collection --</option>'; // Clear existing
  BrandData.forEach((brand) => {
    const option = document.createElement("option");
    option.value = brand.id;
    option.textContent = brand.name;
    container.appendChild(option);
  });
}
function renderCategoryDropdownOptions(CategoryData) {
  const container = document.getElementById("categorySelect");
  container.innerHTML = '<option value="">-- Select Category --</option>'; // Clear existing
  CategoryData.forEach((category) => {
    const option = document.createElement("option");
    option.value = category.id;
    option.textContent = category.name;
    container.appendChild(option);
  });
}
function renderGenderDropdownOptions(GenderData) {
  const container = document.getElementById("genderSelect");
  container.innerHTML = '<option value="">-- Select Gender --</option>'; // Clear existing
  GenderData.forEach((gender) => {
    const option = document.createElement("option");
    option.value = gender.id;
    option.textContent = gender.name;
    container.appendChild(option);
  });
}
function renderStyleDropdownOptions(StylesDatas) {
  const container = document.getElementById("styleOptions");
  if (!container) return;
  container.innerHTML = ''; // Clear existing checkbox options
  // Reset the search input
  const searchInput = document.getElementById("styleSearchInput");
  if (searchInput) {
    searchInput.value = "";
  }
  // Reset the Select All checkbox
  const selectAllCheckbox = document.getElementById("styleSelectAll");
  if (selectAllCheckbox) {
    selectAllCheckbox.checked = false;
  }
  // Reset the dropdown button text
  const btn = document.getElementById("styleDropdownBtn");
  if (btn) {
    btn.title = "--Select Style--";
  }
  const valSpan = document.getElementById("styleDropdownValue");
  if (valSpan) {
    valSpan.textContent = "--Select Style--";
  }
  const display = document.getElementById("selectedStyles");
  if (display) {
    display.textContent = "None";
    display.hidden = true;
  }

  StylesDatas.forEach((style) => {
    const label = document.createElement("label");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = style.StyleId;
    checkbox.dataset.code = style.StyleCode;
    checkbox.className = "style-option-checkbox";
    const textNode = document.createElement("span");
    textNode.textContent = style.StyleCode + "-" + style.Name + (style.FreeField1 && style.FreeField1 !== "null" ? "-" + style.FreeField1 : "");
    label.appendChild(checkbox);
    label.appendChild(textNode);
    container.appendChild(label);
  });
}

// Style Multiselect Dropdown Event Listeners Setup
document.addEventListener("click", (e) => {
  const menu = document.getElementById("styleDropdownMenu");
  const btn = document.getElementById("styleDropdownBtn");
  if (menu && btn) {
    if (e.target === btn) {
      e.stopPropagation();
      const isOpen = menu.classList.toggle("show");
      if (isOpen) {
        btn.classList.add("active");
      } else {
        btn.classList.remove("active");
      }
    } else if (!menu.contains(e.target)) {
      menu.classList.remove("show");
      btn.classList.remove("active");
    }
  }
});

const styleSelectAll = document.getElementById("styleSelectAll");
if (styleSelectAll) {
  styleSelectAll.addEventListener("change", (e) => {
    const isChecked = e.target.checked;
    const checkboxes = document.querySelectorAll(".style-option-checkbox");
    checkboxes.forEach((cb) => {
      const parentLabel = cb.closest("label");
      if (parentLabel && parentLabel.style.display !== "none") {
        cb.checked = isChecked;
      }
    });
    SelectedStylesButton();
  });
}

const styleOptions = document.getElementById("styleOptions");
if (styleOptions) {
  styleOptions.addEventListener("change", (e) => {
    if (e.target && e.target.classList.contains("style-option-checkbox")) {
      SelectedStylesButton();
    }
  });
}

const styleSearchInput = document.getElementById("styleSearchInput");
if (styleSearchInput) {
  styleSearchInput.addEventListener("input", (e) => {
    const query = e.target.value.toLowerCase().trim();
    const labels = document.querySelectorAll("#styleOptions label");

    labels.forEach((label) => {
      const code = label.textContent.toLowerCase();
      if (code.includes(query)) {
        label.style.display = "flex";
      } else {
        label.style.display = "none";
      }
    });
    SelectedStylesButton();
  });
}

function SelectedStylesButton() {
  const checkboxes = document.querySelectorAll(".style-option-checkbox");
  const selectedCodes = [];
  const selectedIds = [];
  let allVisibleChecked = true;
  let hasVisibleCheckboxes = false;
  checkboxes.forEach((cb) => {
    const parentLabel = cb.closest("label");
    const isVisible = parentLabel && parentLabel.style.display !== "none";
    if (cb.checked) {
      selectedCodes.push(cb.dataset.code);
      selectedIds.push(cb.value);
    }
    if (isVisible) {
      hasVisibleCheckboxes = true;
      if (!cb.checked) {
        allVisibleChecked = false;
      }
    }
  });
  const selectAllCheckbox = document.getElementById("styleSelectAll");
  if (selectAllCheckbox) {
    selectAllCheckbox.checked = hasVisibleCheckboxes && allVisibleChecked;
  }

  const btn = document.getElementById("styleDropdownBtn");
  const valSpan = document.getElementById("styleDropdownValue");
  if (btn) {
    if (selectedCodes.length > 0) {
      btn.title = selectedCodes.join(", ");
    } else {
      btn.title = "--Select Style--";
    }
  }
  if (valSpan) {
    if (selectedCodes.length > 0) {
      valSpan.textContent = selectedCodes.join(", ");
    } else {
      valSpan.textContent = "--Select Style--";
    }
  }
  const display = document.getElementById("selectedStyles");
  if (display) {
    display.textContent = selectedIds.length > 0 ? selectedIds.join(",") : "None";
    display.hidden = selectedIds.length === 0;
  }
}

function updateSeasSelectedValues() {
  const select = document.getElementById("seasonSelect");
  const selectedId = select.value;
  const selectedName = select.options[select.selectedIndex]?.textContent;
  select.title = selectedName;
  const idDisplay = document.getElementById("selectedSeasons");
  idDisplay.textContent = selectedId ? selectedId : "None";
  idDisplay.hidden = !selectedId;
  FetchAndShowStylesInDropDown()
}

function updateDivSelectedValues() {
  const select = document.getElementById("divisionSelect");
  const selectedId = select.value;
  const selectedName = select.options[select.selectedIndex]?.textContent;
  select.title = selectedName;
  const idDisplay = document.getElementById("selectedDivisions");
  idDisplay.textContent = selectedId ? selectedId : "None";
  idDisplay.hidden = !selectedId;
  FetchAndShowStylesInDropDown()
}
function updatebrandSelectedValues() {
  const select = document.getElementById("brandSelect");
  const selectedId = select.value;
  const selectedName = select.options[select.selectedIndex]?.textContent;
  select.title = selectedName;
  const idDisplay = document.getElementById("selectedBrands");
  idDisplay.textContent = selectedId ? selectedId : "None";
  idDisplay.hidden = !selectedId;
  FetchAndShowStylesInDropDown()
}
function updatecategorySelectedValues() {
  const select = document.getElementById("categorySelect");
  const selectedId = select.value;
  const selectedName = select.options[select.selectedIndex]?.textContent;
  select.title = selectedName;
  console.log("Selected Category: ", selectedName);
  const idDisplay = document.getElementById("selectedCategories");
  idDisplay.textContent = selectedId ? selectedId : "None";
  idDisplay.hidden = !selectedId;
  FetchAndShowStylesInDropDown()
}
function updategenderSelectedValues() {
  const select = document.getElementById("genderSelect");
  console.log("Selected Gender Element:", select);
  const selectedId = select.value;
  const selectedName = select.options[select.selectedIndex]?.textContent;
  select.title = selectedName;
  console.log("Selected Gender: ", selectedName);
  const idDisplay = document.getElementById("selectedGender");
  idDisplay.textContent = selectedId ? selectedId : "None";
  idDisplay.hidden = !selectedId;
  FetchAndShowStylesInDropDown()
}

document.getElementById("StyleData").addEventListener("click", function () {
  if (!validateSelections()) {
    return;
  }
  document.getElementById("UploadData").style.display = "none";
  const display = document.getElementById("selectedStyles");
  const selectedIdsStr = display ? display.textContent : "None";
  if (!selectedIdsStr || selectedIdsStr === "None") {
    showWarning("Please select at least one style to fetch details.");
    return;
  }
  const selectedIds = selectedIdsStr.split(",");
  console.log("Selected StyleIds:", selectedIds);
  try {
    showLoader();
    const accessToken = document.getElementById("access_token").innerText;
    // Fetch details for all selected styles in parallel
    const fetchPromises = selectedIds.map(id => {
      const styleDetailsUrl = apiURL + `/api/Styles/StyleDetails?styleId=${id}`;
      return fetch(styleDetailsUrl, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json",
        },
      }).then(res => {
        if (!res.ok) {
          throw new Error(`HTTP error! status: ${res.status}`);
        }
        return res.json();
      });
    });

    Promise.all(fetchPromises).then((results) => {
      let combinedStyleDetails = [];
      let rawTrimCategories = [];
      results.forEach((data) => {
        let styleDetails = data.styleData || data.StyleData || [];
        combinedStyleDetails = combinedStyleDetails.concat(styleDetails);
        let trimList = data.trimCategoryList || data.TrimCategoryList || [];
        rawTrimCategories = rawTrimCategories.concat(trimList);
      });
      // DEDUPLICATE THE DICTIONARIES HERE 
      const combinedTrimCategories = rawTrimCategories.filter((item, index, self) =>
        index === self.findIndex((t) => JSON.stringify(t) === JSON.stringify(item))
      );
      console.log("Combined Style Details:", combinedStyleDetails);
      console.log("Unique Trim Categories (Fixed):", combinedTrimCategories);
      if (combinedStyleDetails.length === 0) {
        hideLoader();
        return;
      }
      combinedStyleDetails.forEach((costing) => {
        if (costing.fabricModel) {
          costing.printEmbroideryFabrics = costing.fabricModel.filter((f) => f && (f.materialCategory || "").trim() === "Print & Embroidery");
          costing.fabricModel = costing.fabricModel.filter((f) => f && (f.materialCategory || "").trim() !== "Print & Embroidery");
        } else {
          costing.printEmbroideryFabrics = [];
        }
      });
      StyleResponseData = combinedStyleDetails;
      try {
        localStorage.setItem("StyleResponseData", JSON.stringify(combinedStyleDetails));
      } catch (e) {
        console.error("Failed to save StyleResponseData to localStorage:", e);
      }
      if (typeof Office !== "undefined" && Office.context?.document?.settings) {
        try {
          Office.context.document.settings.set("StyleResponseData", combinedStyleDetails);
          Office.context.document.settings.saveAsync();
        } catch (e) {
          console.error("Failed to save StyleResponseData to Office settings:", e);
        }
      }
      WriteStyleDatainSheet(StyleResponseData, combinedTrimCategories);
    })
      .catch((error) => {
        hideLoader();
        console.error("Error fetching style details:", error);
      });
  } catch (err) {
    console.error("Error in StyleData click handler for fetch data: ", err);
  }
});

function getTrimSubFields(section) {
  return [                     /*Trim sub fields for all Costing model at Trims & Accessories sections*/
    { label: "Number", prop: "trimNumber" },
    { label: "Name", prop: "trimName" },
    { label: "Negotiated Price", prop: "costPrice" },
    { label: "Consumption / Dzn", prop: "quantity" },
    { label: "Consumption & Price UOM", prop: "priceUOM" },
  ];
}
function buildFieldMap(data) {
  const isKnit = data?.[0]?.costingModel === "Knit / Sweater Costing V1";
  const maxFabrics = Math.max(...data.map((c) => c.fabricModel?.length || 0));

  const staticFields = [            /* Static fields for all Costing model at Overview label*/
    { label: "Style Number", key: "styleNumber" },
    { label: "Style Name", key: "styleName" },
    { label: "Collection Line", key: "collectionLine" },
    { label: "Season", key: "season" },
    { label: "Category", key: "category" },
    { label: "Selected Buyer", key: "selectedBuyer" },
    { label: "Style Colorway", key: "styleColorway" },
    { label: "Costing Model", key: "costingModel" },
    { label: "SupplierId", key: "supplierId" },
    { label: "Supplier CountryId", key: "supplierCountryId" },
    { label: "Costing Column Sequence", key: "costingColumnSequence" },
    { label: "BOM Version", key: "bomVersion" },
    { label: "BOM Notes", key: "bomNotes" },
  ];
  let mainFabricFlatFields;
  if (isKnit) {
    mainFabricFlatFields = [
      { label: "Supplier", key: "mainFabricSupplier" },
      { label: "Supplier Country", key: "mainFabricSupplierCountry" },
      { label: "Supplier Reference", key: "mainFabricReference" },
      { label: "Material Number", key: "mainFabricFabricName" },
      { label: "Composition", key: "mainFabricComposition" },
      { label: "Gauge", key: "mainFabricConstruction" },
      { label: "Minimum Order Quantity", key: "mainFabricMOQ" },
      { label: "MOQ UOM", key: "mainFabricMOQUOM" },
      { label: "Finish", key: "mainFabricFinish" },
      { label: "Placement", key: "mainFabricPlacement" },
      { label: "Weight", key: "mainFabricWeight" },
      { label: "Weight UOM", key: "mainFabricWeightUOM" },
      { label: "Width", key: "mainFabricWidth" },
      { label: "Width UOM", key: "mainFabricWidthUOM" },
      { label: "Negotiated Price", key: "mainFabricCostPrice" },
      { label: "Consumption / Dzn", key: "mainFabricQuantity" },
      { label: "Consumption & Price UOM", key: "mainFabricUOM" },
      { label: "Total Shell Fabric Cost", key: null, isFormula: "totalShellFabricCost" },
    ];
  } else {
    mainFabricFlatFields = [
      { label: "Supplier", key: "mainFabricSupplier" },
      { label: "Supplier Country", key: "mainFabricSupplierCountry" },
      { label: "Supplier Reference", key: "mainFabricReference" },
      { label: "Material Number", key: "mainFabricFabricName" },
      { label: "Composition", key: "mainFabricComposition" },
      { label: "Construction", key: "mainFabricConstruction" },
      { label: "Minimum Order Quantity", key: "mainFabricQuantity" },
      { label: "MOQ UOM", key: "mainFabricUOM" },
      { label: "Finish", key: "mainFabricFinish" },
      { label: "Placement", key: "mainFabricPlacement" },
      { label: "Weight", key: "mainFabricWeight" },
      { label: "Weight UOM", key: "mainFabricWeightUOM" },
      { label: "Width", key: "mainFabricWidth" },
      { label: "Width UOM", key: "mainFabricWidthUOM" },
      { label: "Negotiated Price", key: "mainFabricCostPrice" },
      { label: "Consumption / Dzn", key: "mainFabricQuantity" },
      { label: "Consumption & Price UOM", key: "mainFabricUOM" },
      { label: "Total Shell Fabric Cost", key: null, isFormula: "totalShellFabricCost" },
    ];
  }
  const fabricFields = [];
  if (isKnit) {
    for (let i = 0; i < maxFabrics + 1; i++) {
      const prefix = `Yarn ${i + 2}`;
      fabricFields.push(
        { label: `${prefix} - Material Number`, key: `fabricModel[${i}].fabricNumber` },
        { label: `${prefix} - Material Name`, key: `fabricModel[${i}].fabricName` },
        { label: `${prefix} - Width`, key: `fabricModel[${i}].width` },
        { label: `${prefix} - Width UOM`, key: `fabricModel[${i}].widthUOM` },
        { label: `${prefix} - Negotiated Price`, key: `fabricModel[${i}].costPrice` },
        { label: `${prefix} - Consumption / Dzn`, key: `fabricModel[${i}].quantity` },
        { label: `${prefix} - Consumption & Price UOM`, key: `fabricModel[${i}].uom` },
      );
    }
  } else {
    for (let i = 0; i < maxFabrics + 1; i++) {
      const prefix = `Fabric ${i + 2}`;
      fabricFields.push(
        { label: `${prefix} - Material Number`, key: `fabricModel[${i}].fabricNumber` },
        { label: `${prefix} - Material Name`, key: `fabricModel[${i}].fabricName` },
        { label: `${prefix} - Weight`, key: `fabricModel[${i}].weight` },
        { label: `${prefix} - Weight UOM`, key: `fabricModel[${i}].weightUOM` },
        { label: `${prefix} - Negotiated Price`, key: `fabricModel[${i}].costPrice` },
        { label: `${prefix} - Consumption / Dzn`, key: `fabricModel[${i}].quantity` },
        { label: `${prefix} - Consumption & Price UOM`, key: `fabricModel[${i}].uom` },
      );
    }
  }
  const trimFields = [];
  const categoryOrderMap = new Map();
  // Define default categories/ Group Heading that should always appear on the sheet
  const defaultCategories = [
    "Sewing / Labeling Accessories",
    "Finishing / Packing Accessories",
    "Sewing / Labeling Trims",
    "Finishing / Packing Trims",
    "Embellishments",
    "Print & Embroidery",
    "Wash"
  ];
  defaultCategories.forEach((cat) => {
    categoryOrderMap.set(cat, data.map(() => []));
  });

  for (let colIdx = 0; colIdx < data.length; colIdx++) {
    const trims = data[colIdx]?.trimModel || [];
    trims.forEach((trim, trimIndex) => {
      let cat = (trim?.trimMaterialCategory || "Uncategorized").trim();
      // Normalize category names to map to the default categories
      if (cat === "Sewing/Labeling Accessories") cat = "Sewing / Labeling Accessories";
      if (cat === "Finishing/Packing Accessories") cat = "Finishing / Packing Accessories";
      if (cat === "Sewing/Labeling Trims") cat = "Sewing / Labeling Trims";
      if (cat === "Finishing/Packing Trims") cat = "Finishing / Packing Trims";
      if (cat === "Embellishment") cat = "Embellishments";

      if (!categoryOrderMap.has(cat)) {
        categoryOrderMap.set(cat, data.map(() => []));
      }
      categoryOrderMap.get(cat)[colIdx].push({ source: "trim", index: trimIndex });
    });

    const peFabrics = data[colIdx]?.printEmbroideryFabrics || [];
    peFabrics.forEach((fabric, fabricIndex) => {
      const cat = "Print & Embroidery";
      if (!categoryOrderMap.has(cat)) {
        categoryOrderMap.set(cat, data.map(() => []));
      }
      categoryOrderMap.get(cat)[colIdx].push({ source: "fabric", index: fabricIndex });
    });
  }

  const getTrimSection = (cat) => {
    if (isKnit) {
      const map = {
        "Sewing / Labeling Accessories": "4 - Sewing/Labeling Accessories",
        "Sewing/Labeling Accessories": "4 - Sewing/Labeling Accessories",
        "Finishing / Packing Accessories": "5 - Finishing/Packing Accessories",
        "Finishing/Packing Accessories": "5 - Finishing/Packing Accessories",
        "Sewing / Labeling Trims": "6 - Sewing/Labeling Trims",
        "Sewing/Labeling Trims": "6 - Sewing/Labeling Trims",
        "Finishing / Packing Trims": "7 - Finishing/Packing Trims",
        "Finishing/Packing Trims": "7 - Finishing/Packing Trims",
        "Embellishments": "8 - Embellishments",
        "Print & Embroidery": "9 - Print & Embroidery Information",
        "Wash": "10 - Wash Information",
      };
      return map[cat] || "4 - Sewing/Labeling Accessories";
    } else {
      const map = {
        "Sewing / Labeling Trims": "4 - Sewing/Labeling Trims",
        "Sewing/Labeling Trims": "4 - Sewing/Labeling Trims",
        "Sewing / Labeling Accessories": "5 - Sewing/Labeling Accessories",
        "Sewing/Labeling Accessories": "5 - Sewing/Labeling Accessories",
        "Finishing / Packing Trims": "6 - Finishing/Packing Trims",
        "Finishing/Packing Trims": "6 - Finishing/Packing Trims",
        "Finishing / Packing Accessories": "7 - Finishing/Packing Accessories",
        "Finishing/Packing Accessories": "7 - Finishing/Packing Accessories",
        "Embellishments": "8 - Embellishments",
        "Print & Embroidery": "9 - Print & Embroidery Information",
        "Wash": "10 - Wash Information",
      };
      return map[cat] || "5 - Sewing/Labeling Accessories";
    }
  };

  let orderedCategories;

  if (isKnit) {
    orderedCategories = [
      "Sewing / Labeling Accessories",
      "Sewing/Labeling Accessories",
      "Finishing / Packing Accessories",
      "Finishing/Packing Accessories",
      "Sewing / Labeling Trims",
      "Sewing/Labeling Trims",
      "Finishing / Packing Trims",
      "Finishing/Packing Trims",
      "Embellishments",
      "Print & Embroidery",
      "Wash",
    ];
  } else {
    orderedCategories = [
      "Sewing / Labeling Trims",
      "Sewing/Labeling Trims",
      "Sewing / Labeling Accessories",
      "Sewing/Labeling Accessories",
      "Finishing / Packing Trims",
      "Finishing/Packing Trims",
      "Finishing / Packing Accessories",
      "Finishing/Packing Accessories",
      "Embellishments",
      "Print & Embroidery",
      "Wash",
    ];
  }

  const sortedCategories = Array.from(categoryOrderMap.keys()).sort((a, b) => {
    let idxA = orderedCategories.indexOf(a);
    let idxB = orderedCategories.indexOf(b);
    if (idxA === -1) idxA = 999;
    if (idxB === -1) idxB = 999;
    return idxA - idxB;
  });

  for (const cat of sortedCategories) {
    const trimIndicesPerCol = categoryOrderMap.get(cat);
    const section = getTrimSection(cat);
    const maxInGroup = Math.max(...trimIndicesPerCol.map((arr) => arr.length), 0);
    const subFields = getTrimSubFields(section);

    for (let slot = 0; slot < maxInGroup + 1; slot++) {
      subFields.forEach(({ label, prop }) => {
        trimFields.push({
          label,
          key: null,
          section,
          _trimGroupCategory: cat,
          _trimGroupSlot: slot,
          _trimProp: prop,
          _trimIndicesPerCol: trimIndicesPerCol,
        });
      });
    }
  }
  let costFields;

  if (isKnit) {
    costFields = [
      // Knit / Sweater Costing fields
      { label: "BOM Yarn Cost", key: "bomFabricCost", section: "11 - BOM Cost" },
      { label: "BOM Sewing/Labeling Accessories Cost", key: "bomSewingLabelingAccessoriesCost", section: "11 - BOM Cost", _bomSectionRef: "4 - Sewing/Labeling Accessories" },
      { label: "BOM Finishing/Packing Accessories Cost", key: "bomLabelandPackagingCost", section: "11 - BOM Cost", _bomSectionRef: "5 - Finishing/Packing Accessories" },
      { label: "BOM Sewing/Labeling Trims Cost", key: "bomSewingtrimcost", section: "11 - BOM Cost", _bomSectionRef: "6 - Sewing/Labeling Trims" },
      { label: "BOM Finishing/Packing Trims Cost", key: "bomPackagintrimcost", section: "11 - BOM Cost", _bomSectionRef: "7 - Finishing/Packing Trims" },
      { label: "BOM Embellishments Cost", key: "bomTrimCost", section: "11 - BOM Cost", _bomSectionRef: "8 - Embellishments" },
      { label: "Print / Embroidery Cost", key: "printembroidarycost", section: "11 - BOM Cost", _bomSectionRef: "9 - Print & Embroidery Information" },
      { label: "Wash Cost", key: "washcost", section: "11 - BOM Cost", _bomSectionRef: "10 - Wash Information" },
      { label: "Total Raw Material Cost", key: "totalRawMaterialCost", section: "11 - BOM Cost" },

      { label: "CM from Library", key: "cM_WashAndPrintCost", section: "12 - Operations Cost" },
      { label: "Efficiency %", key: "efficiency", section: "12 - Operations Cost" },
      { label: "Calculated CM", key: null, section: "12 - Operations Cost", isFormula: "calculatedCm" },
      { label: "Negotiated CM - Input", key: "negotiated", section: "12 - Operations Cost" },
      { label: "Inland Import Charges", key: "inlandImportCharges", section: "12 - Operations Cost" },
      { label: "Inland Export Charges", key: "inlandExportCharges", section: "12 - Operations Cost" },
      { label: "Commercial", key: "commercial", section: "12 - Operations Cost" },
      { label: "Lab Test", key: "labTest", section: "12 - Operations Cost" },
      { label: "Additional MUP/FIN Cost", key: "additionalMUPorFINCost", section: "12 - Operations Cost" },
      { label: "Certification Cost", key: "certificationCost", section: "12 - Operations Cost" },

      { label: "Order Quantity", key: "orderQty", section: "13 - Fixed Cost" },
      { label: "Inspection Cost", key: "inspectionCost", section: "13 - Fixed Cost" },
      { label: "Inspection Cost per Pc", key: null, section: "13 - Fixed Cost", isFormula: "inspectionCostPerPc" },

      { label: "Total Cost", key: "totalCost", section: "14 - Total Cost" },
      { label: "Total Cost per Pc", key: "totalCostPerPc", section: "14 - Total Cost" },

      { label: "Fixed Margin %", key: "fixedMargin", section: "15 - FOB Price" },
      { label: "Calculated FOB Price", key: "calculatedFOBPricePerPc", section: "15 - FOB Price" },
      { label: "Proposed FOB Price", key: "proposedFOBPrice", section: "15 - FOB Price" },

      { label: "Calculated Margin %", key: "calculatedMargin", section: "16 - Margin" },
      { label: "KAM's Euro Conversion Rate", key: "kaMsEuroConversationRate", section: "16 - Margin" },
      { label: "FOB Price in Euro", key: "marginValueInEuro", section: "16 - Margin", isFormula: "fobPriceInEuro" },
      { label: "Buyers Target Price in USD", key: "buyerTargetprice", section: "16 - Margin" },
      { label: "Buyers Target Price in Euro", key: "revisedTargetPrice", section: "16 - Margin", isFormula: "buyersTargetPriceInEuro" },

      { label: "Transport", key: "transport", section: "17 - DDP Price" },
      { label: "Taux", key: "taux", section: "17 - DDP Price" },
      { label: "DDP Price", key: "additionalMUPorFINCost", section: "17 - DDP Price", isFormula: "ddpPrice" },
      { label: "Total Calculated FOB Value", key: null, section: "17 - DDP Price", isFormula: "totalCalculatedFobValue" },
      { label: "Update Y/N", key: "update", section: "18 - Action" },
    ];
  } else {
    costFields = [
      // Woven Costing fields
      { label: "BOM Fabric Cost", key: "bomFabricCost", section: "11 - BOM Cost" },
      { label: "BOM Sewing/Labeling Accessories Cost", key: "bomSewingLabelingAccessoriesCost", section: "11 - BOM Cost", _bomSectionRef: "5 - Sewing/Labeling Accessories" },
      { label: "BOM Finishing/Packing Accessories Cost", key: "bomLabelandPackagingCost", section: "11 - BOM Cost", _bomSectionRef: "7 - Finishing/Packing Accessories" },
      { label: "BOM Sewing/Labeling Trims Cost", key: "bomSewingtrimcost", section: "11 - BOM Cost", _bomSectionRef: "4 - Sewing/Labeling Trims" },
      { label: "BOM Finishing/Packing Trims Cost", key: "bomPackagintrimcost", section: "11 - BOM Cost", _bomSectionRef: "6 - Finishing/Packing Trims" },
      { label: "BOM Embellishments Cost", key: "bomTrimCost", section: "11 - BOM Cost", _bomSectionRef: "8 - Embellishments" },
      { label: "Print / Embroidery Cost", key: "printembroidarycost", section: "11 - BOM Cost", _bomSectionRef: "9 - Print & Embroidery Information" },
      { label: "Wash Cost", key: "washcost", section: "11 - BOM Cost", _bomSectionRef: "10 - Wash Information" },
      { label: "Total Raw Material Cost", key: "totalRawMaterialCost", section: "11 - BOM Cost" },

      { label: "CM from Library", key: "cM_WashAndPrintCost", section: "12 - Operations Cost" },
      { label: "Efficiency %", key: "efficiency", section: "12 - Operations Cost" },
      { label: "Calculated CM", key: null, section: "12 - Operations Cost", isFormula: "calculatedCm" },
      { label: "Negotiated CM - Input", key: "negotiated", section: "12 - Operations Cost" },
      { label: "Lab Test", key: "labTest", section: "12 - Operations Cost" },
      { label: "Certification Cost", key: "certificationCost", section: "12 - Operations Cost" },

      { label: "Order Quantity", key: "orderQty", section: "13 - Fixed Cost" },
      { label: "Inspection Cost", key: "inspectionCost", section: "13 - Fixed Cost" },
      { label: "Inspection Cost per Pc", key: null, section: "13 - Fixed Cost", isFormula: "inspectionCostPerPc" },

      { label: "Total Cost", key: "totalCost", section: "14 - Total Cost" },
      { label: "Total Cost per Pc", key: "totalCostPerPc", section: "14 - Total Cost" },

      { label: "Fixed Cost and Wastage %", key: "fixedCostAndWastege", section: "15 - FOB Price" },
      { label: "Fixed Margin %", key: "fixedMargin", section: "15 - FOB Price" },
      { label: "Fixed Rebate %", key: "fixedRebate", section: "15 - FOB Price" },
      { label: "Calculated FOB Price", key: "calculatedFOBPricePerPc", section: "15 - FOB Price" },
      { label: "Proposed FOB Price", key: "proposedFOBPrice", section: "15 - FOB Price" },

      { label: "Calculated Margin %", key: "calculatedMargin", section: "16 - Margin" },
      { label: "KAM's Euro Conversion Rate", key: "kaMsEuroConversationRate", section: "16 - Margin" },
      { label: "FOB Price in Euro", key: "marginValueInEuro", section: "16 - Margin", isFormula: "fobPriceInEuro" },
      { label: "Buyers Target Price in USD", key: "buyerTargetprice", section: "16 - Margin" },
      { label: "Buyers Target Price in Euro", key: "revisedTargetPrice", section: "16 - Margin", isFormula: "buyersTargetPriceInEuro" },

      { label: "Transport", key: "transport", section: "17 - DDP Price" },
      { label: "Taux", key: "taux", section: "17 - DDP Price" },
      { label: "DDP Price", key: "additionalMUPorFINCost", section: "17 - DDP Price", isFormula: "ddpPrice" },
      { label: "Total Calculated FOB Value", key: null, section: "17 - DDP Price", isFormula: "totalCalculatedFobValue" },
      { label: "Update Y/N", key: "update", section: "18 - Action" },
    ];
  }

  const sec2Label = isKnit ? "2 - Yarn Information" : "2 - Shell Fabric Information";
  const sec3Label = isKnit ? "3 - Other Yarn Information" : "3 - Other Fabric Information";

  return [
    ...staticFields.map((f) => ({ ...f, section: "1 - Style Information" })),
    ...mainFabricFlatFields.map((f) => ({ ...f, section: sec2Label })),
    ...fabricFields.map((f) => ({ ...f, section: sec3Label })),
    ...trimFields,
    ...costFields,
  ];
}
/* Data Printing in Sheet logic */
async function WriteStyleDatainSheet(data, trimCategories) {
  document.getElementById("UploadData").style.display = "none";
  try {
    showLoader();
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      function getNestedValue(obj, path) {
        try {
          return path.split(/[\.\[\]]+/).filter(Boolean).reduce((acc, part) => {
            if (acc === null || acc === undefined) return "";
            return acc[isNaN(part) ? part : Number(part)];
          }, obj) ?? "";
        } catch (err) {
          console.error("Error resolving nested value:", err);
          return "";
        }
      }
      const sectionColorMap = {
        "1 - Style Information": "#F2F2F2",
        "2 - Shell Fabric Information": "#F7ECE6",
        "2 - Yarn Information": "#F7ECE6",
        "3 - Other Fabric Information": "#FCE4D6",
        "3 - Other Yarn Information": "#FCE4D6",
        "4 - Sewing/Labeling Trims": "#FCE4D6",
        "4 - Sewing/Labeling Accessories": "#FCE4D6",
        "5 - Sewing/Labeling Accessories": "#FCE4D6",
        "5 - Finishing/Packing Accessories": "#FCE4D6",
        "6 - Finishing/Packing Trims": "#FCE4D6",
        "6 - Sewing/Labeling Trims": "#FCE4D6",
        "7 - Finishing/Packing Accessories": "#FCE4D6",
        "7 - Finishing/Packing Trims": "#FCE4D6",
        "8 - Embellishments": "#FCE4D6",
        "9 - Print & Embroidery Information": "#FCE4D6",
        "10 - Wash Information": "#FCE4D6",
        "11 - BOM Cost": "#F4B084",
        "12 - Operations Cost": "#76C44B",
        "13 - Fixed Cost": "#B3429A",
        "14 - Total Cost": "#9BC2E6",
        "15 - FOB Price": "#9BC2E6",
        "16 - Margin": "#DDEBF7",
        "17 - DDP Price": "#DDEBF7",
        "18 - Action": "#FFFFFF",
      };

      const maxFabrics = Math.max(...data.map((c) => c.fabricModel?.length || 0));
      console.log(`Dynamic: ${maxFabrics} fabrics`);

      const fieldMap = buildFieldMap(data);

      // ── Build header row ──────────────────────────────────────────────────
      const styleCounts = {};
      const headerRow = [
        "Costing Sections",
        "Costing Elements",
        ...data.map((costing) => {
          const styleNum = costing.styleNumber || "";
          if (styleCounts[styleNum] === undefined) {
            styleCounts[styleNum] = 0;
          }
          styleCounts[styleNum]++;
          return `Costing ${styleCounts[styleNum]}`;
        }),
      ];

      const dataRows = fieldMap.map(
        ({ section, label, key, isFormula, _bomSectionRef, _trimGroupCategory, _trimGroupSlot, _trimProp, _trimIndicesPerCol }) => {
          const displayLabel = (label || "").replace(/^(Fabric|Yarn)\s+\d+\s*-\s*/i, "");
          return [
            section ?? "",
            displayLabel,
            ...data.map((costing, colIdx) => {
              if (isFormula || _bomSectionRef) {
                return "";
              }
              //  Grouped trim resolution
              if (_trimGroupCategory !== undefined) {
                const indices = _trimIndicesPerCol?.[colIdx] || [];
                const itemRef = indices[_trimGroupSlot];
                if (itemRef === undefined || itemRef === null) {
                  return "";
                }
                if (typeof itemRef === "number") {
                  const trim = costing?.trimModel?.[itemRef];
                  if (!trim) {
                    return "";
                  }
                  return trim[_trimProp] ?? "";
                } else if (itemRef.source === "trim") {
                  const trim = costing?.trimModel?.[itemRef.index];
                  if (!trim) {
                    return "";
                  }
                  return trim[_trimProp] ?? "";
                } else if (itemRef.source === "fabric") {
                  const fabric = costing?.printEmbroideryFabrics?.[itemRef.index];
                  if (!fabric) {
                    return "";
                  }
                  let fabricProp = _trimProp;
                  if (_trimProp === "trimNumber") fabricProp = "fabricNumber";
                  if (_trimProp === "trimName") fabricProp = "fabricName";
                  if (_trimProp === "currency") fabricProp = "uom";
                  return fabric[fabricProp] ?? "";
                }
                return "";
              }
              // Normal field resolution
              if (key === "" || !key) return "";
              return getNestedValue(costing, key) ?? "";
            }),
          ];
        }
      );
      const allRows = [headerRow, ...dataRows];
      const totalRows = allRows.length;
      const totalCols = headerRow.length;

      // ── Write to sheet ────────────────────────────────────────────────────
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (!usedRange.isNullObject) {
        usedRange.clear(Excel.ClearApplyTo.contents);
        usedRange.clear(Excel.ClearApplyTo.formats);
        try {
          usedRange.ungroup(Excel.GroupOption.byRows);
        } catch (e) {
          console.error("Error ungrouping range:" + e.message);
          console.log("Error ungrouping range:" + e.stack);
          console.warn("Error ungrouping range:" + e.message);
          // ignore if no groups to ungroup
        }
      }

      const range = sheet.getRangeByIndexes(0, 0, totalRows, totalCols);
      console.log("AllRows sections:", allRows.map((r) => r[0]));
      range.values = allRows;

      // ── Row index helpers ─────────────────────────────────────────────────
      const labelsList = fieldMap.map((f) => f.label);
      const getSheetRow = (label) => labelsList.indexOf(label) + 2;

      const isKnit = data?.[0]?.costingModel === "Knit / Sweater Costing V1";

      const rowTotalShellFabricCost = getSheetRow("Total Shell Fabric Cost");
      const rowBomFabricCost = getSheetRow(isKnit ? "BOM Yarn Cost" : "BOM Fabric Cost");
      const rowBomSewLabelAccCost = getSheetRow("BOM Sewing/Labeling Accessories Cost");
      const rowBomFinPackAccCost = getSheetRow("BOM Finishing/Packing Accessories Cost");
      const rowBomSewLabelTrimsCost = getSheetRow("BOM Sewing/Labeling Trims Cost");
      const rowBomFinPackTrimsCost = getSheetRow("BOM Finishing/Packing Trims Cost");
      const rowBomEmbellishmentsCost = getSheetRow("BOM Embellishments Cost");
      const rowPrintEmbroideryCost = getSheetRow("Print / Embroidery Cost");
      const rowWashCost = getSheetRow("Wash Cost");
      const rowTotalRawMaterial = getSheetRow("Total Raw Material Cost");

      const rowCmFromLibrary = getSheetRow("CM from Library");
      const rowEfficiency = getSheetRow("Efficiency %");
      const rowCalculatedCm = getSheetRow("Calculated CM");
      const rowNegotiatedCmInput = getSheetRow("Negotiated CM - Input");
      const rowLabTest = getSheetRow("Lab Test");
      const rowCertification = getSheetRow("Certification Cost");

      const rowInlandImport = isKnit ? getSheetRow("Inland Import Charges") : 0;
      const rowInlandExport = isKnit ? getSheetRow("Inland Export Charges") : 0;
      const rowCommercial = isKnit ? getSheetRow("Commercial") : 0;
      const rowAdditionalMup = isKnit ? getSheetRow("Additional MUP/FIN Cost") : 0;

      const rowOrderQuantity = getSheetRow("Order Quantity");
      const rowInspectionCost = getSheetRow("Inspection Cost");
      const rowInspectionCostPerPc = getSheetRow("Inspection Cost per Pc");

      const rowTotalCost = getSheetRow("Total Cost");
      const rowTotalCostPerPc = getSheetRow("Total Cost per Pc");

      const rowFixedCostAndWastage = isKnit ? 0 : getSheetRow("Fixed Cost and Wastage %");
      const rowFixedMargin = getSheetRow("Fixed Margin %");
      const rowFixedRebate = isKnit ? 0 : getSheetRow("Fixed Rebate %");
      const rowCalculatedFobPrice = getSheetRow("Calculated FOB Price");
      const rowProposedFobPrice = getSheetRow("Proposed FOB Price");

      const rowCalculatedMargin = getSheetRow("Calculated Margin %");
      const rowKamsEuroConvRate = getSheetRow("KAM's Euro Conversion Rate");
      const rowFobPriceInEuro = getSheetRow("FOB Price in Euro");
      const rowBuyersTargetPriceUSD = getSheetRow("Buyers Target Price in USD");
      const rowBuyersTargetPriceEuro = getSheetRow("Buyers Target Price in Euro");

      const rowTransport = getSheetRow("Transport");
      const rowTaux = getSheetRow("Taux");
      const rowDdpPrice = getSheetRow("DDP Price");
      const rowTotalCalculatedFobValue = getSheetRow("Total Calculated FOB Value");
      const mainFabricCostRow = getSheetRow("Negotiated Price");
      const mainFabricQuantityRow = getSheetRow("Consumption / Dzn");

      // ── Total Shell Fabric Cost formula (main fabric only) ────────────────
      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const col = String.fromCharCode(65 + colIdx);
        sheet.getRangeByIndexes(rowTotalShellFabricCost - 1, colIdx, 1, 1).formulas = [[
          `=IFERROR(${col}${mainFabricCostRow}*${col}${mainFabricQuantityRow},0)`
        ]];
      }

      // ── BOM Fabric Cost formula (main fabric + all dynamic fabrics) ───────
      if (isKnit) {
        for (let colIdx = 2; colIdx < totalCols; colIdx++) {
          const col = String.fromCharCode(65 + colIdx);
          let parts = [`${col}${rowTotalShellFabricCost}`];
          for (let i = 0; i < maxFabrics + 1; i++) {
            const prefix = `Yarn ${i + 2}`;
            const qtyRow = getSheetRow(`${prefix} - Consumption / Dzn`);
            const priceRow = getSheetRow(`${prefix} - Negotiated Price`);
            parts.push(`IFERROR(${col}${qtyRow}*${col}${priceRow},0)`);
          }
          const formula = `=` + parts.join("+");
          sheet.getRangeByIndexes(rowBomFabricCost - 1, colIdx, 1, 1).formulas = [[formula]];
        }
      } else {
        const fabricPairs = [{ fabricCostRow: mainFabricCostRow, fabricQuantityRow: mainFabricQuantityRow }];
        for (let i = 0; i < maxFabrics + 1; i++) {
          const prefix = `Fabric ${i + 2}`;
          fabricPairs.push({
            fabricCostRow: getSheetRow(`${prefix} - Negotiated Price`),
            fabricQuantityRow: getSheetRow(`${prefix} - Consumption / Dzn`),
          });
        }
        for (let colIdx = 2; colIdx < totalCols; colIdx++) {
          const col = String.fromCharCode(65 + colIdx);
          const formula = fabricPairs.length > 0
            ? `=` + fabricPairs.map(({ fabricCostRow, fabricQuantityRow }) =>
              `IFERROR(${col}${fabricCostRow}*${col}${fabricQuantityRow},0)`).join("+")
            : `=0`;
          sheet.getRangeByIndexes(rowBomFabricCost - 1, colIdx, 1, 1).formulas = [[formula]];
        }
      }

      const groupRowMap = new Map();
      fieldMap.forEach((f, idx) => {
        if (f._trimGroupCategory === undefined) return;
        const slotKey = `${f._trimGroupCategory}__${f._trimGroupSlot}`;
        if (!groupRowMap.has(slotKey)) groupRowMap.set(slotKey, { section: f.section });
        const rec = groupRowMap.get(slotKey);
        if (f._trimProp === "costPrice") rec.costRow = idx + 2;
        if (f._trimProp === "quantity") rec.qtyRow = idx + 2;
      });

      const sectionTrimPairs = {}; // section -> [{costRow, qtyRow}, ...]
      groupRowMap.forEach(({ costRow, qtyRow, section }) => {
        if (!costRow || !qtyRow) return;
        if (!sectionTrimPairs[section]) sectionTrimPairs[section] = [];
        sectionTrimPairs[section].push({ costRow, qtyRow });
      });

      function writeSectionCostFormula(targetRow, sourceSection) {
        const pairs = sectionTrimPairs[sourceSection] || [];
        for (let colIdx = 2; colIdx < totalCols; colIdx++) {
          const col = String.fromCharCode(65 + colIdx);
          const formula = pairs.length > 0
            ? `=` + pairs.map(({ costRow, qtyRow }) => `IFERROR(${col}${costRow}*${col}${qtyRow},0)`).join("+")
            : `=0`;
          sheet.getRangeByIndexes(targetRow - 1, colIdx, 1, 1).formulas = [[formula]];
        }
      }
      // ── Per-section BOM cost formulas ──────────────────────────────────────
      if (isKnit) {
        writeSectionCostFormula(rowBomSewLabelAccCost, "4 - Sewing/Labeling Accessories");
        writeSectionCostFormula(rowBomFinPackAccCost, "5 - Finishing/Packing Accessories");
        writeSectionCostFormula(rowBomSewLabelTrimsCost, "6 - Sewing/Labeling Trims");
        writeSectionCostFormula(rowBomFinPackTrimsCost, "7 - Finishing/Packing Trims");
      } else {
        writeSectionCostFormula(rowBomSewLabelAccCost, "5 - Sewing/Labeling Accessories");
        writeSectionCostFormula(rowBomFinPackAccCost, "7 - Finishing/Packing Accessories");
        writeSectionCostFormula(rowBomSewLabelTrimsCost, "4 - Sewing/Labeling Trims");
        writeSectionCostFormula(rowBomFinPackTrimsCost, "6 - Finishing/Packing Trims");
      }
      writeSectionCostFormula(rowBomEmbellishmentsCost, "8 - Embellishments");
      writeSectionCostFormula(rowPrintEmbroideryCost, "9 - Print & Embroidery Information");
      writeSectionCostFormula(rowWashCost, "10 - Wash Information");

      // ── Total Raw Material Cost formula ───────────────────────────────────
      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const col = String.fromCharCode(65 + colIdx);
        sheet.getRangeByIndexes(rowTotalRawMaterial - 1, colIdx, 1, 1).formulas = [[
          `=SUM(${col}${rowBomFabricCost}:${col}${rowWashCost})`
        ]];
      }

      // ── Calculated CM formula ─────────────────────────────────────────────
      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const col = String.fromCharCode(65 + colIdx);
        sheet.getRangeByIndexes(rowCalculatedCm - 1, colIdx, 1, 1).formulas = [[
          `=IFERROR(IF(${col}${rowEfficiency}=0,${col}${rowCmFromLibrary},${col}${rowCmFromLibrary}/(${col}${rowEfficiency}/100)),0)`
        ]];
      }

      // ── Inspection Cost per Pc formula ────────────────────────────────────
      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const col = String.fromCharCode(65 + colIdx);
        sheet.getRangeByIndexes(rowInspectionCostPerPc - 1, colIdx, 1, 1).formulas = [[
          `=IFERROR(IF(${col}${rowOrderQuantity}=0,0,${col}${rowInspectionCost}/(${col}${rowOrderQuantity}/12)),0)`
        ]];
      }

      // ── Total Cost formula ────────────────────────────────────────────────
      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const col = String.fromCharCode(65 + colIdx);
        const formula = isKnit
          ? `=IFERROR(IF(${col}${rowNegotiatedCmInput}=0,${col}${rowInspectionCostPerPc}+${col}${rowCertification}+${col}${rowAdditionalMup}+${col}${rowLabTest}+${col}${rowCommercial}+${col}${rowInlandExport}+${col}${rowInlandImport}+${col}${rowCalculatedCm}+${col}${rowTotalRawMaterial},${col}${rowInspectionCostPerPc}+${col}${rowCertification}+${col}${rowAdditionalMup}+${col}${rowNegotiatedCmInput}+${col}${rowTotalRawMaterial}),0)`
          : `=IFERROR(IF(${col}${rowNegotiatedCmInput}<>0,${col}${rowNegotiatedCmInput},${col}${rowCalculatedCm})+${col}${rowTotalRawMaterial}+${col}${rowLabTest}+${col}${rowCertification}+${col}${rowInspectionCostPerPc},0)`;
        sheet.getRangeByIndexes(rowTotalCost - 1, colIdx, 1, 1).formulas = [[formula]];
      }

      // ── Total Cost per Pc formula ─────────────────────────────────────────
      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const col = String.fromCharCode(65 + colIdx);
        sheet.getRangeByIndexes(rowTotalCostPerPc - 1, colIdx, 1, 1).formulas = [[
          `=IFERROR(${col}${rowTotalCost}/12,0)`
        ]];
      }

      // ── Calculated FOB Price formula ──────────────────────────────────────
      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const col = String.fromCharCode(65 + colIdx);
        if (isKnit) {
          sheet.getRangeByIndexes(rowCalculatedFobPrice - 1, colIdx, 1, 1).formulas = [[
            `=IFERROR(IF(${col}${rowFixedMargin}=0,${col}${rowTotalCostPerPc},${col}${rowTotalCostPerPc}/(1-(${col}${rowFixedMargin}/100))),0)`
          ]];
        } else {
          const formula1 = `IF(${col}${rowFixedCostAndWastage}=0,0,IF(${col}${rowFixedMargin}=0,0,IF(${col}${rowFixedRebate}=0,0,IF(${col}${rowTotalCostPerPc}=0,0,IF((1-(${col}${rowFixedCostAndWastage}+${col}${rowFixedMargin}))=0,0,IF((1-(${col}${rowFixedRebate}/100))=0,0,${col}${rowTotalCostPerPc}/(1-(${col}${rowFixedCostAndWastage}/100+${col}${rowFixedMargin}/100))/(1-(${col}${rowFixedRebate}/100))))))))`;
          const formula2 = `IF(${col}${rowTotalCostPerPc}=0,0,IF(${col}${rowFixedMargin}=0,0,(${col}${rowTotalCostPerPc}/(1-(${col}${rowFixedMargin}/100)))))`;
          sheet.getRangeByIndexes(rowCalculatedFobPrice - 1, colIdx, 1, 1).formulas = [[
            `=IFERROR(IF(${col}${rowFixedCostAndWastage}<>0,${formula1},${formula2}),0)`
          ]];
        }
      }

      // ── Calculated Margin % formula ───────────────────────────────────────
      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const col = String.fromCharCode(65 + colIdx);
        sheet.getRangeByIndexes(rowCalculatedMargin - 1, colIdx, 1, 1).formulas = [[
          `=IFERROR(IF(${col}${rowProposedFobPrice}=0,0,(${col}${rowProposedFobPrice}-${col}${rowTotalCostPerPc})/${col}${rowProposedFobPrice}*100),0)`
        ]];
      }

      // ── FOB Price in Euro formula ─────────────────────────────────────────
      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const col = String.fromCharCode(65 + colIdx);
        sheet.getRangeByIndexes(rowFobPriceInEuro - 1, colIdx, 1, 1).formulas = [[
          `=IFERROR(${col}${rowProposedFobPrice}*${col}${rowKamsEuroConvRate},0)`
        ]];
      }

      // ── Buyers Target Price in Euro formula ───────────────────────────────
      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const col = String.fromCharCode(65 + colIdx);
        sheet.getRangeByIndexes(rowBuyersTargetPriceEuro - 1, colIdx, 1, 1).formulas = [[
          `=IFERROR(${col}${rowBuyersTargetPriceUSD}*${col}${rowKamsEuroConvRate},0)`
        ]];
      }

      // ── DDP Price formula ─────────────────────────────────────────────────
      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const col = String.fromCharCode(65 + colIdx);
        const formula = `=IFERROR(${col}${rowTransport}+(${col}${rowProposedFobPrice}/${col}${rowTaux}),0)`;
        sheet.getRangeByIndexes(rowDdpPrice - 1, colIdx, 1, 1).formulas = [[formula]];
      }

      // ── Total Calculated FOB Value formula ────────────────────────────────
      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const col = String.fromCharCode(65 + colIdx);
        sheet.getRangeByIndexes(rowTotalCalculatedFobValue - 1, colIdx, 1, 1).formulas = [[
          `=IFERROR(${col}${rowCalculatedFobPrice}*${col}${rowOrderQuantity},0)`
        ]];
      }

      // ── Number formatting ─────────────────────────────────────────────────
      for (let r = 1; r < totalRows; r++) {
        const section = allRows[r][0];
        const label = allRows[r][1] || "";
        const labelLower = label.toLowerCase();
        // Skip Overview section (text)
        if (section.startsWith("1 - ")) continue;
        // Skip Action section (text/update options)
        if (section.startsWith("18 - ")) continue;
        // Skip non-numeric fields
        if (
          labelLower.includes("number") ||
          labelLower.includes("name") ||
          labelLower.includes("uom") ||
          labelLower.includes("supplier") ||
          labelLower.includes("reference") ||
          labelLower.includes("composition") ||
          labelLower.includes("construction") ||
          labelLower.includes("gauge") ||
          labelLower.includes("finish") ||
          labelLower.includes("placement") ||
          labelLower.includes("currency") ||
          labelLower.includes("action") ||
          labelLower.includes("update") ||
          labelLower.includes("consumption")
        ) {
          continue;
        }

        // Apply number format (#,##0.00) to numeric rows for columns index 2 onwards
        sheet.getRangeByIndexes(r, 2, 1, totalCols - 2).numberFormat = [["#,##0.00"]];
      }

      // ── STEP 9: Disable Gridlines ─────────────────────────────────────────
      sheet.showGridlines = false;

      // ── STEP 10a: Full-sheet light border grid ────────────────────────────
      const fullSheetRange = sheet.getRange("A1:Z500");
      const allBorderTypes = [
        Excel.BorderIndex.edgeTop,
        Excel.BorderIndex.edgeBottom,
        Excel.BorderIndex.edgeLeft,
        Excel.BorderIndex.edgeRight,
        Excel.BorderIndex.insideHorizontal,
        Excel.BorderIndex.insideVertical,
      ];
      allBorderTypes.forEach((borderType) => {
        const border = fullSheetRange.format.borders.getItem(borderType);
        border.style = Excel.BorderLineStyle.continuous;
        border.weight = Excel.BorderWeight.thin;
        border.color = "#D3D3D3";
      });

      // ── STEP 10b: Dark borders on used data range ─────────────────────────
      const verticalRange = sheet.getRangeByIndexes(0, 0, totalRows, totalCols);
      verticalRange.format.horizontalAlignment = "Left";
      [
        Excel.BorderIndex.edgeLeft,
        Excel.BorderIndex.edgeRight,
        Excel.BorderIndex.insideVertical,
      ].forEach((borderType) => {
        const border = verticalRange.format.borders.getItem(borderType);
        border.style = Excel.BorderLineStyle.continuous;
        border.weight = Excel.BorderWeight.thin;
        border.color = "#000000";
      });

      // ── STEP 11: Section color fill + Bottom Border ───────────────────────
      let colorStartRow = 1;
      while (colorStartRow < allRows.length) {
        const section = allRows[colorStartRow][0];
        let colorEndRow = colorStartRow;
        while (
          colorEndRow + 1 < allRows.length &&
          allRows[colorEndRow + 1][0] === section
        ) {
          colorEndRow++;
        }
        const sectionColor = sectionColorMap[section] || "#FFFFFF";
        // Fill entire section
        sheet.getRangeByIndexes(
          colorStartRow,
          0,
          colorEndRow - colorStartRow + 1,
          totalCols
        ).format.fill.color = sectionColor;

        // Bottom Border
        var sectionBottomBorder = sheet
          .getRangeByIndexes(colorEndRow, 0, 1, totalCols)
          .format.borders.getItem(Excel.BorderIndex.edgeBottom);

        sectionBottomBorder.style = Excel.BorderLineStyle.continuous;
        sectionBottomBorder.weight = Excel.BorderWeight.medium;
        sectionBottomBorder.color = "#000000";

        colorStartRow = colorEndRow + 1;
      }

      // ── STEP 12: Header row formatting (unchanged) ────────────────────────
      const headerRange = sheet.getRangeByIndexes(0, 0, 1, totalCols);
      headerRange.format.font.bold = true;
      headerRange.format.font.color = "#020303";
      headerRange.format.fill.color = "#FFFFFF";
      const headerBottomBorder = headerRange.format.borders.getItem(Excel.BorderIndex.edgeBottom);
      headerBottomBorder.style = Excel.BorderLineStyle.continuous;
      headerBottomBorder.weight = Excel.BorderWeight.medium;
      headerBottomBorder.color = "#000000";

      // ── STEP 13: Thin sub-item dividers ────────────────────────────────────
      for (let r = 1; r < allRows.length - 1; r++) {
        const section = allRows[r][0];
        const label = allRows[r][1] || "";

        const isFabricOrTrimSection =
          section.startsWith("3 - ") || section.startsWith("4 - ") ||
          section.startsWith("5 - ") || section.startsWith("6 - ") ||
          section.startsWith("7 - ") || section.startsWith("8 - ") ||
          section.startsWith("9 - ") || section.startsWith("10 - ");

        if (isFabricOrTrimSection) {
          const isLastField = label === "Consumption & Price UOM" || label === "Consumption/Purchase UOM";

          if (isLastField) {
            const bottomBorder = sheet
              .getRangeByIndexes(r, 0, 1, totalCols)
              .format.borders.getItem(Excel.BorderIndex.edgeBottom);
            bottomBorder.style = Excel.BorderLineStyle.continuous;
            bottomBorder.weight = Excel.BorderWeight.thin;
            bottomBorder.color = "#000000";
          }
        }
      }

      // ── STEP 14: Row bolding + white input fills ──────────────────────────
      function isInputRow(section, label) {
        const secLower = section.toLowerCase();
        const labelLower = label.toLowerCase();

        if (
          secLower.startsWith("2 - ") ||
          secLower.startsWith("3 - ") ||
          secLower.startsWith("4 - ") ||
          secLower.startsWith("5 - ") ||
          secLower.startsWith("6 - ") ||
          secLower.startsWith("7 - ") ||
          secLower.startsWith("8 - ") ||
          secLower.startsWith("9 - ") ||
          secLower.startsWith("10 - ")
        ) {
          return labelLower.includes("negotiated price") || labelLower.includes("consumption / dzn");
        }

        const inputRowLabels = [
          "efficiency %",
          "negotiated cm - input",
          "inland import charges",
          "inland export charges",
          "commercial",
          "lab test",
          "additional mup/fin cost",
          "certification cost",
          "order quantity",
          "inspection cost",
          "fixed cost and wastage %",
          "fixed margin %",
          "fixed rebate %",
          "proposed fob price",
          "kam's euro conversion rate",
          "buyers target price in usd",
          "transport",
          "taux",
          "update y/n",
          "action for update"
        ];
        return inputRowLabels.includes(labelLower) || secLower.startsWith("18 - ");
      }

      for (let r = 1; r < allRows.length; r++) {
        const section = allRows[r][0];
        const label = allRows[r][1] || "";
        const dataRange = sheet.getRangeByIndexes(r, 2, 1, totalCols - 2);

        // Column A (Section) and Column B (Label) are always bold
        sheet.getRangeByIndexes(r, 0, 1, 2).format.font.bold = true;

        if (isInputRow(section, label)) {
          if (section.startsWith("18 - ") || label.toLowerCase() === "action for update") {
            sheet.getRangeByIndexes(r, 1, 1, totalCols - 1).format.fill.color = "#FFFFFF";
            sheet.getRangeByIndexes(r, 1, 1, totalCols - 1).format.font.bold = false;
          } else {
            dataRange.format.fill.color = "#FFFFFF";
            dataRange.format.font.bold = false;
          }
        } else {
          const shouldValueBeBold = label.toLowerCase().includes("total") || /^(11|12|13|14|15|16|17) -/.test(section);
          dataRange.format.font.bold = shouldValueBeBold;
        }
      }
      // ── STEP 15: Action row alignment ─────────────────────────────────────
      sheet.getRangeByIndexes(totalRows - 1, 1, 1, totalCols - 1)
        .format.horizontalAlignment = Excel.HorizontalAlignment.center;

      // ── STEP 16: Bold column A labels ─────────────────────────────────────
      sheet.getRangeByIndexes(1, 0, totalRows - 1, 1).format.font.bold = true;

      // ── STEP 17: Merge + align section cells in column A ──────────────────
      let mergeStartRow = 1;
      while (mergeStartRow < allRows.length) {
        const section = allRows[mergeStartRow][0];
        let mergeEndRow = mergeStartRow;
        while (mergeEndRow + 1 < allRows.length && allRows[mergeEndRow + 1][0] === section) {
          mergeEndRow++;
        }
        if (mergeEndRow > mergeStartRow) {
          const r = sheet.getRangeByIndexes(mergeStartRow, 0, mergeEndRow - mergeStartRow + 1, 1);
          r.merge(false);
          r.format.verticalAlignment = Excel.VerticalAlignment.top;
          r.format.horizontalAlignment = Excel.HorizontalAlignment.left;
          r.format.font.bold = true;
        }
        mergeStartRow = mergeEndRow + 1;
      }
      const sectionRanges = [];
      let groupStartRow = 1;
      while (groupStartRow < allRows.length) {
        const section = allRows[groupStartRow][0];
        let groupEndRow = groupStartRow;
        while (groupEndRow + 1 < allRows.length && allRows[groupEndRow + 1][0] === section) {
          groupEndRow++;
        }
        if (groupEndRow > groupStartRow) {
          sectionRanges.push({ start: groupStartRow, count: groupEndRow - groupStartRow });
        }
        groupStartRow = groupEndRow + 1;
      }

      for (const { start, count } of sectionRanges) {
        try {
          const detailRange = sheet.getRangeByIndexes(start, 0, count, totalCols);
          detailRange.group(Excel.GroupOption.byRows);
          await context.sync(); // commit this group before starting the next one
        } catch (groupError) {
          console.error(`Failed to group rows ${start + 1} to ${start + count}:`, groupError);
        }
      }

      // ── STEP 18: Autofit columns ──────────────────────────────────────────
      range.format.autofitColumns();

      await context.sync();
      hideLoader();
      console.log(`Sheet written: ${totalRows} rows x ${totalCols} cols | Fabrics: ${maxFabrics}`);
    }).then(async () => {
      await toggleButtonVisibility();
    }).catch((error) => {
      hideLoader();
      console.error("Error writing style data to sheet:", error);
    });
  } catch (err) {
    hideLoader();
    console.error("Error in WriteStyleDatainSheet:", err);
  } finally {
    hideLoader();
  }
}
document.getElementById("clearButton").addEventListener("click", () => {
  try {
    localStorage.removeItem("StyleResponseData");
  } catch (e) {
    console.error("Failed to remove StyleResponseData from localStorage:", e);
  }
  if (typeof Office !== "undefined" && Office.context?.document?.settings) {
    try {
      Office.context.document.settings.remove("StyleResponseData");
      Office.context.document.settings.saveAsync();
    } catch (e) {
      console.error("Failed to remove StyleResponseData from Office settings:", e);
    }
  }
  StyleResponseData = null;

  Excel.run(async (context) => {
    document.getElementById("UploadData").style.display = "none";
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    // Clear values and formatting
    usedRange.clear(Excel.ClearApplyTo.contents);
    usedRange.clear(Excel.ClearApplyTo.formats);
    sheet.showGridlines = true;
    await context.sync();
    // Delete all shapes
    const shapes = sheet.shapes;
    shapes.load("items");
    await context.sync();
    shapes.items.forEach(shape => shape.delete());

    try {
      usedRange.ungroup(Excel.GroupOption.byRows);
    } catch (e) {
      // ignore if no groups
    }

    await context.sync();
    console.log("Sheet cleared.");
  }).catch((error) => {
    console.error("Error clearing sheet:", error);
  });
});


async function onSheetChanged(event) {
  console.log("onSheetChanged event.address:", event.address);
  const m = String(event.address).match(/([A-Z]+)\d+/i);
  const colLetter = m ? m[1].toUpperCase() : "";
  const rowNumber = m ? parseInt(m[0].replace(/[A-Z]/gi, "")) : 0;
  console.log("Changed col:", colLetter, "row:", rowNumber);
  if (!colLetter) return;
  // Debounce rapid changes
  if (sheetChangeDebounce) clearTimeout(sheetChangeDebounce);
  sheetChangeDebounce = setTimeout(async () => {
    if (isRunning) {
      console.log("onSheetChanged skipped, another operation is running.");
      return;
    }
    isRunning = true;
    try {
      // ── Toggle button — single call only  ──────────────────────────────────
      await toggleButtonVisibility();
      // ── Char limit enforcement ────────────────────────────────────────────────
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load(["isNullObject", "rowCount"]);
        await context.sync();
        if (usedRange.isNullObject) return;
        const colARange = sheet.getRangeByIndexes(0, 0, usedRange.rowCount, 1);
        colARange.load("values");
        await context.sync();
        const rowLabels = colARange.values.map(r => String(r[0]).trim());
        const changedRowLabel = (rowNumber > 0 && rowLabels[rowNumber - 1]) ? String(rowLabels[rowNumber - 1]).trim() : "";
        console.log("Changed row label:", changedRowLabel);
        const notesLabels = [
          "Inland Import Charges",
          "Inland Export Charges",
          "Commercial",
          "Lab Test",
          "Inspection Cost",
          "Certification Cost",
          "KAM's Euro Conversation Rate",
        ];
        const freeFieldLabels = [];

        if (changedRowLabel && notesLabels.some(l => l.toLowerCase() === changedRowLabel.toLowerCase())) {
          await enforceCharLimit(event.address, 1024, colLetter);
        } else if (changedRowLabel && freeFieldLabels.length > 0 &&
          freeFieldLabels.some(l => l.toLowerCase() === changedRowLabel.toLowerCase())) {
          await enforceCharLimit(event.address, 255, colLetter);
        }
      });
    } catch (error) {
      if (error.message && error.message.indexOf("cell-editing mode") !== -1) {
        console.warn("onSheetChanged: Excel is in cell-editing mode. Operation deferred.");
      } else {
        console.error("onSheetChanged error:", error);
      }
    } finally {
      isRunning = false;
    }
  }, 200); // debounce delay
}

function findSheetRowIndex(allData, section, fieldLabel, slotIndex = 0) {
  const targetSec = String(section || "").trim().toLowerCase();
  const expectedDisplayLabel = String(fieldLabel || "")
    .replace(/^(Fabric|Yarn)\s+\d+\s*-\s*/i, "")
    .trim()
    .toLowerCase();

  let currentSection = "";
  let occurrenceCount = 0;

  for (let idx = 0; idx < allData.length; idx++) {
    const secColValue = String(allData[idx][0] || "").trim();
    if (secColValue !== "") {
      currentSection = secColValue.toLowerCase();
    }

    if (currentSection === targetSec) {
      const sheetLabel = String(allData[idx][1] || "").trim().toLowerCase();
      if (sheetLabel === expectedDisplayLabel) {
        if (occurrenceCount === slotIndex) {
          return idx;
        }
        occurrenceCount++;
      }
    }
  }
  return -1;
}

function mapToCostingUpdateRequest(costingName, colIndex, costingObj) {
  // Map fabricModel to FabricModelUpdate, prepending the main fabric/yarn information
  const fabricModeData = [];

  // 1. Add main fabric/yarn details if fabric number exists
  if (costingObj.mainFabricFabricName && String(costingObj.mainFabricFabricName).trim() !== "") {
    fabricModeData.push({
      "Fabric Number": String(costingObj.mainFabricFabricName).trim(),
      "Cost Price": String(costingObj.mainFabricCostPrice !== undefined && costingObj.mainFabricCostPrice !== null ? costingObj.mainFabricCostPrice : "0"),
      "Quantity": String(costingObj.mainFabricQuantity !== undefined && costingObj.mainFabricQuantity !== null ? costingObj.mainFabricQuantity : "0")
    });
  }
  // 2. Add other fabric/yarn details
  if (Array.isArray(costingObj.fabricModel)) {
    costingObj.fabricModel.forEach(f => {
      if (f) {
        fabricModeData.push({
          "Fabric Number": f.fabricNumber || "",
          "Cost Price": String(f.costPrice !== undefined && f.costPrice !== null ? f.costPrice : "0"),
          "Quantity": String(f.quantity !== undefined && f.quantity !== null ? f.quantity : "0")
        });
      }
    });
  }
  // Map trimModel to TrimModelUpdate
  const trimModeData = Array.isArray(costingObj.trimModel)
    ? costingObj.trimModel.map(t => ({
      "Trim Number": t.trimNumber || "",
      "Cost Price": String(t.costPrice !== undefined && t.costPrice !== null ? t.costPrice : "0"),
      "Quantity": String(t.quantity !== undefined && t.quantity !== null ? t.quantity : "0")
    }))
    : [];

  return {
    "Costing Name": costingName,
    "ColIndex": String(colIndex),
    "StyleNumber": costingObj.styleNumber || "",
    "Season": costingObj.season || "",
    // "Main Fabric Name": costingObj.mainFabricName || "",
    // "Main fabric Currency": costingObj.mainFabricCurrency || "",
    "SupplierId": String(costingObj.supplierId !== undefined && costingObj.supplierId !== null ? costingObj.supplierId : ""),
    "BOM Version": costingObj.bomVersion !== undefined && costingObj.bomVersion !== null ? String(costingObj.bomVersion) : "",
    // "BOMVersion": costingObj.bomVersion !== undefined && costingObj.bomVersion !== null ? String(costingObj.bomVersion) : "",
    // "Main Fabric Quantity": String(costingObj.mainFabricQuantity !== undefined && costingObj.mainFabricQuantity !== null ? costingObj.mainFabricQuantity : "0"),
    "Additional MUP / FIN Cost": String(costingObj.additionalMUPorFINCost !== undefined && costingObj.additionalMUPorFINCost !== null ? costingObj.additionalMUPorFINCost : "0"),
    "Negotiated CM - Input": String(costingObj.negotiated !== undefined && costingObj.negotiated !== null ? costingObj.negotiated : "0"),
    "Inland Import Charge": String(costingObj.inlandImportCharges !== undefined && costingObj.inlandImportCharges !== null ? costingObj.inlandImportCharges : "0"),
    "Inland Export Charge": String(costingObj.inlandExportCharges !== undefined && costingObj.inlandExportCharges !== null ? costingObj.inlandExportCharges : "0"),
    "Commercial": String(costingObj.commercial !== undefined && costingObj.commercial !== null ? costingObj.commercial : "0"),
    "Lab Test": String(costingObj.labTest !== undefined && costingObj.labTest !== null ? costingObj.labTest : "0"),
    "Inspection Cost": String(costingObj.inspectionCost !== undefined && costingObj.inspectionCost !== null ? costingObj.inspectionCost : "0"),
    "Certification Cost": String(costingObj.certificationCost !== undefined && costingObj.certificationCost !== null ? costingObj.certificationCost : "0"),
    // "KAMs Euro Conversation Rate": String(costingObj.kaMsEuroConversationRate !== undefined && costingObj.kaMsEuroConversationRate !== null ? costingObj.kaMsEuroConversationRate : "0"),
    // "Consumption In Dozen Per Pc": String(costingObj.consumptionInDozenPerPc !== undefined && costingObj.consumptionInDozenPerPc !== null ? costingObj.consumptionInDozenPerPc : "0"),
    "Target Price": String(costingObj.buyerTargetprice !== undefined && costingObj.buyerTargetprice !== null ? costingObj.buyerTargetprice : (costingObj.buyerTargetprice !== undefined && costingObj.buyerTargetprice !== null ? costingObj.buyerTargetprice : "0")),
    "Revised Target Price": String(costingObj.revisedTargetPrice !== undefined && costingObj.revisedTargetPrice !== null ? costingObj.revisedTargetPrice : "0"),
    "Efficiency": String(costingObj.efficiency !== undefined && costingObj.efficiency !== null ? costingObj.efficiency : "0"),
    "Order Quantity": String(costingObj.orderQty !== undefined && costingObj.orderQty !== null ? costingObj.orderQty : "0"),
    "Proposed FOB Price": String(costingObj.proposedFOBPrice !== undefined && costingObj.proposedFOBPrice !== null ? costingObj.proposedFOBPrice : "0"),
    "Transport": String(costingObj.transport !== undefined && costingObj.transport !== null ? costingObj.transport : "0"),
    "Taux": String(costingObj.taux !== undefined && costingObj.taux !== null ? costingObj.taux : "0"),
    "update": String(costingObj.update !== undefined && costingObj.update !== null ? costingObj.update : ""),
    "FabricModel": fabricModeData,
    "TrimModel": trimModeData
  };
}

document.getElementById("UploadData").addEventListener("click", async function getExcelData() {
  showLoader();
  const accessToken = document.getElementById("access_token").innerText;
  try {
    let allData = null;
    let totalCols = 0;
    let totalRows = 0;
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load(["isNullObject", "columnCount", "rowCount"]);
      await context.sync();
      if (usedRange.isNullObject || usedRange.columnCount < 3) {
        return;
      }
      totalCols = usedRange.columnCount;
      totalRows = usedRange.rowCount;
      const dataRange = sheet.getRangeByIndexes(0, 0, totalRows, totalCols);
      dataRange.load("values");
      await context.sync();
      allData = dataRange.values;
    });

    if (!allData) {
      showWarning("No data found in sheet.");
      hideLoader();
      return;
    }

    if (!StyleResponseData || StyleResponseData.length === 0) {
      try {
        if (typeof localStorage !== "undefined") {
          const cachedData = localStorage.getItem("StyleResponseData");
          if (cachedData) {
            StyleResponseData = JSON.parse(cachedData);
          }
        }
      } catch (e) {
        console.error("Failed to load StyleResponseData from localStorage in getExcelData:", e);
      }
      if (!StyleResponseData && typeof Office !== "undefined" && Office.context?.document?.settings) {
        try {
          const savedData = Office.context.document.settings.get("StyleResponseData");
          if (savedData) {
            StyleResponseData = typeof savedData === "string" ? JSON.parse(savedData) : savedData;
          }
        } catch (e) {
          console.error("Failed to load StyleResponseData from Office settings in getExcelData:", e);
        }
      }
    }

    if (!StyleResponseData || StyleResponseData.length === 0) {
      lastDebugLog = "";
      await ensureStyleResponseData(allData, totalCols);
    }

    if (!StyleResponseData || StyleResponseData.length === 0) {
      showWarning("No original style data found. Please select Season, Department, Collection Line, Gender, and Category dropdowns to automatically resolve style details, or fetch style details again." + (lastDebugLog ? "\n\nDebug Details:\n" + lastDebugLog : ""));
      hideLoader();
      return;
    }
    const fieldMap = buildFieldMap(StyleResponseData);

    const rowLabels = allData.map(row => String(row[0]).trim());
    const actionRowIndex = rowLabels.findIndex((label, idx) => {
      const valA = label.toLowerCase();
      const valB = String(allData[idx][1] || "").trim().toLowerCase();
      return valA === "action for update" || valA === "update y/n" || valA === "18 - action" ||
        valB === "action for update" || valB === "update y/n" || valB === "18 - action" ||
        valA.includes("action") || valA.includes("update") ||
        valB.includes("action") || valB.includes("update");
    });

    if (actionRowIndex === -1) {
      showWarning("Action for Update or Update Y/N row not found.");
      hideLoader();
      return;
    }
    console.log("Action row index:", actionRowIndex);
    const actionRowValues = allData[actionRowIndex];
    const flaggedColIndexes = [];
    for (let colIdx = 2; colIdx < totalCols; colIdx++) {
      const val = actionRowValues[colIdx];
      if (val === 1 || val === "1" || String(val).trim().toUpperCase() === "Y") {
        flaggedColIndexes.push(colIdx);
      }
    }
    console.log("Flagged col indexes:", flaggedColIndexes);
    if (flaggedColIndexes.length === 0) {
      showWarning("No costing columns flagged for upload. Enter Y in Update Y/N row.");
      hideLoader();
      return;
    }

    const costingPayload = flaggedColIndexes.map((colIdx) => {
      const costingName = String(allData[0][colIdx]).trim();
      const costingObj = (StyleResponseData && StyleResponseData[colIdx - 2])
        ? JSON.parse(JSON.stringify(StyleResponseData[colIdx - 2]))
        : { fabricModel: [], trimModel: [], printEmbroideryFabrics: [] };

      fieldMap.forEach((field, fieldIdx) => {
        let slotIndex = 0;
        if (field.key && field.key.startsWith("fabricModel[")) {
          const match = field.key.match(/fabricModel\[(\d+)\]/);
          if (match) {
            slotIndex = parseInt(match[1]);
          }
        } else if (field._trimGroupCategory !== undefined) {
          slotIndex = field._trimGroupSlot || 0;
        }

        const rowIdx = findSheetRowIndex(allData, field.section, field.label, slotIndex);
        if (rowIdx === -1) {
          return;
        }
        const cellValue = allData[rowIdx][colIdx];

        if (field.key) {
          if (field.key.startsWith("fabricModel[")) {
            const match = field.key.match(/fabricModel\[(\d+)\]\.(.+)/);
            if (match) {
              const fabIdx = parseInt(match[1]);
              const prop = match[2];
              if (!costingObj.fabricModel) costingObj.fabricModel = [];
              if (!costingObj.fabricModel[fabIdx]) costingObj.fabricModel[fabIdx] = {};

              let val = cellValue;
              const numProps = ["costPrice", "quantity", "weight"];
              if (numProps.includes(prop)) {
                val = (cellValue === "" || cellValue === null || cellValue === undefined) ? 0 : Number(cellValue);
              }
              costingObj.fabricModel[fabIdx][prop] = val;
            }
          } else {
            let val = cellValue;
            const numKeys = [
              "cM_WashAndPrintCost", "negotiated", "labTest", "certificationCost", "inspectionCost",
              "fixedCostAndWastege", "fixedMargin", "fixedRebate", "calculatedFOBPricePerPc", "proposedFOBPrice",
              "kaMsEuroConversationRate", "marginValueInEuro", "targetPrice", "buyerTargetprice", "revisedTargetPrice",
              "additionalMUPorFINCost", "inlandImportCharges", "inlandExportCharges", "commercial",
              "consumptionInDozenPerPc", "totalRawMaterialCost", "mainFabricCostPrice", "mainFabricQuantity",
              "bomFabricCost", "bomSewingLabelingAccessoriesCost", "bomLabelandPackagingCost", "bomSewingtrimcost",
              "bomPackagintrimcost", "bomTrimCost", "printembroidarycost", "washcost",
              "efficiency", "orderQty", "transport", "taux"
            ];
            if (numKeys.includes(field.key)) {
              val = (cellValue === "" || cellValue === null || cellValue === undefined) ? 0 : Number(cellValue);
            }
            costingObj[field.key] = val;
          }
        } else if (field._trimGroupCategory !== undefined) {
          const colIndices = field._trimIndicesPerCol?.[colIdx - 2] || [];
          const itemRef = colIndices[field._trimGroupSlot];

          if (itemRef !== undefined && itemRef !== null) {
            if (typeof itemRef === "number") {
              const trimIndex = itemRef;
              if (!costingObj.trimModel) costingObj.trimModel = [];
              if (!costingObj.trimModel[trimIndex]) {
                const origTrim = StyleResponseData?.[colIdx - 2]?.trimModel?.[trimIndex] || {};
                costingObj.trimModel[trimIndex] = JSON.parse(JSON.stringify(origTrim));
              }

              let val = cellValue;
              if (field._trimProp === "costPrice" || field._trimProp === "quantity") {
                val = (cellValue === "" || cellValue === null || cellValue === undefined) ? 0 : Number(cellValue);
              }
              costingObj.trimModel[trimIndex][field._trimProp] = val;
              costingObj.trimModel[trimIndex].trimMaterialCategory = field._trimGroupCategory;
            } else if (itemRef.source === "trim") {
              const trimIndex = itemRef.index;
              if (!costingObj.trimModel) costingObj.trimModel = [];
              if (!costingObj.trimModel[trimIndex]) {
                const origTrim = StyleResponseData?.[colIdx - 2]?.trimModel?.[trimIndex] || {};
                costingObj.trimModel[trimIndex] = JSON.parse(JSON.stringify(origTrim));
              }

              let val = cellValue;
              if (field._trimProp === "costPrice" || field._trimProp === "quantity") {
                val = (cellValue === "" || cellValue === null || cellValue === undefined) ? 0 : Number(cellValue);
              }
              costingObj.trimModel[trimIndex][field._trimProp] = val;
              costingObj.trimModel[trimIndex].trimMaterialCategory = field._trimGroupCategory;
            } else if (itemRef.source === "fabric") {
              const fabricIndex = itemRef.index;
              if (!costingObj.printEmbroideryFabrics) costingObj.printEmbroideryFabrics = [];
              if (!costingObj.printEmbroideryFabrics[fabricIndex]) {
                const origFabric = StyleResponseData?.[colIdx - 2]?.printEmbroideryFabrics?.[fabricIndex] || {};
                costingObj.printEmbroideryFabrics[fabricIndex] = JSON.parse(JSON.stringify(origFabric));
              }

              let val = cellValue;
              if (field._trimProp === "costPrice" || field._trimProp === "quantity") {
                val = (cellValue === "" || cellValue === null || cellValue === undefined) ? 0 : Number(cellValue);
              }

              let fabricProp = field._trimProp;
              if (field._trimProp === "trimNumber") fabricProp = "fabricNumber";
              if (field._trimProp === "trimName") fabricProp = "fabricName";
              if (field._trimProp === "currency") fabricProp = "uom";

              costingObj.printEmbroideryFabrics[fabricIndex][fabricProp] = val;
              costingObj.printEmbroideryFabrics[fabricIndex].materialCategory = field._trimGroupCategory;
            }
          } else {
            // It is an extra slot. Resolve dynamically using cache on costingObj.
            if (!costingObj._extraSlotTrimIndexCache) {
              costingObj._extraSlotTrimIndexCache = {};
            }
            const cacheKey = `${field._trimGroupCategory}__${field._trimGroupSlot}`;
            if (costingObj._extraSlotTrimIndexCache[cacheKey] === undefined) {
              const originalTrims = StyleResponseData?.[colIdx - 2]?.trimModel || [];
              const newIdx = costingObj.trimModel ? costingObj.trimModel.length : originalTrims.length;
              costingObj._extraSlotTrimIndexCache[cacheKey] = newIdx;
              if (!costingObj.trimModel) costingObj.trimModel = [];
              costingObj.trimModel[newIdx] = {
                trimMaterialCategory: field._trimGroupCategory,
              };
            }
            const trimIndex = costingObj._extraSlotTrimIndexCache[cacheKey];

            if (trimIndex !== undefined && trimIndex !== null) {
              if (!costingObj.trimModel) costingObj.trimModel = [];
              let val = cellValue;
              if (field._trimProp === "costPrice" || field._trimProp === "quantity") {
                val = (cellValue === "" || cellValue === null || cellValue === undefined) ? 0 : Number(cellValue);
              }
              costingObj.trimModel[trimIndex][field._trimProp] = val;
              costingObj.trimModel[trimIndex].trimMaterialCategory = field._trimGroupCategory;
            }
          }
        }
      });

      // Combine fabricModel and printEmbroideryFabrics back together
      let finalFabrics = [];
      if (costingObj.fabricModel) {
        finalFabrics = finalFabrics.concat(costingObj.fabricModel);
      }
      if (costingObj.printEmbroideryFabrics) {
        finalFabrics = finalFabrics.concat(costingObj.printEmbroideryFabrics);
      }
      costingObj.fabricModel = finalFabrics;
      delete costingObj.printEmbroideryFabrics;
      // Filter out completely unpopulated/empty dynamic fabric/yarn & trim objects
      if (costingObj.fabricModel) {
        costingObj.fabricModel = costingObj.fabricModel.filter(
          (f) => f && (f.fabricNumber || f.costPrice || f.quantity)
        );
      }
      if (costingObj.trimModel) {
        costingObj.trimModel = costingObj.trimModel.filter(
          (t) => t && (t.trimNumber || t.costPrice || t.quantity)
        );
      }
      delete costingObj._extraSlotTrimIndexCache;
      return mapToCostingUpdateRequest(costingName, colIdx, costingObj);
    });
    console.log("Structured costingPayload to upload:", JSON.stringify(costingPayload, null, 2));
    const updateUrl = apiURL + "/api/UpdateStyle";
    const response = await fetch(updateUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify(costingPayload),
    });
    console.log("After fetch Json Payload", costingPayload);
    console.log("After fetch response:", response);
    if (!response.ok) {
      const errorResponse = await response.text();
      console.error("Server error details:", errorResponse);
      throw new Error(`HTTP error! status: ${response.status} - Details: ${errorResponse}`);
    }
    const result = await response.json();
    console.log("API response:", result);
    const successCols = result.successCols || [];
    const failureCols = result.failureCols || [];

    const colStatusMap = {};
    flaggedColIndexes.forEach((colIdx) => {
      colStatusMap[colIdx] = {
        isSuccess: successCols.includes(colIdx) || successCols.includes(String(colIdx)),
        isFailure: failureCols.includes(colIdx) || failureCols.includes(String(colIdx)),
        messages: []
      };
    });

    const rootMessages = [];
    if (result.messages) {
      if (Array.isArray(result.messages)) rootMessages.push(...result.messages);
      else rootMessages.push(result.messages);
    } else if (result.message) {
      if (Array.isArray(result.message)) rootMessages.push(...result.message);
      else rootMessages.push(result.message);
    }

    const updateDataList1 = result.updateData || [];
    updateDataList1.forEach((item) => {
      const successCols = item.successCols || [];
      const failledCols = item.failureCols || [];

      let msgs = [];
      if (item.messages) {
        msgs = Array.isArray(item.messages) ? item.messages : [item.messages];
      } else if (item.message) {
        msgs = Array.isArray(item.message) ? item.message : [item.message];
      }

      successCols.forEach((c) => {
        const colIdx = parseInt(c);
        if (!isNaN(colIdx) && colStatusMap[colIdx]) {
          colStatusMap[colIdx].isSuccess = true;
          colStatusMap[colIdx].messages.push(...msgs);
        }
      });

      failledCols.forEach((c) => {
        const colIdx = parseInt(c);
        if (!isNaN(colIdx) && colStatusMap[colIdx]) {
          colStatusMap[colIdx].isFailure = true;
          colStatusMap[colIdx].messages.push(...msgs);
        }
      });
    });

    // Fallback to root messages if no specific messages exist for a column
    flaggedColIndexes.forEach((colIdx) => {
      if (colStatusMap[colIdx].messages.length === 0 && rootMessages.length > 0) {
        colStatusMap[colIdx].messages = [...rootMessages];
      }
    });

    const nonMatchedList = [];

    // 1. Extract from nested updateData
    const updateDataList = result.updateData || [];
    updateDataList.forEach((item) => {
      const sCols = item.successCols || [];
      const fCols = item.failureCols || [];
      const itemCols = [];
      sCols.forEach((c) => { if (c !== "") itemCols.push(Number(c)); });
      fCols.forEach((c) => { if (c !== "") itemCols.push(Number(c)); });

      const nonMatchedLines = item.nonMatchedBomLine || item.nonMatchedBomLines || [];
      if (Array.isArray(nonMatchedLines)) {
        nonMatchedLines.forEach((nonMatch) => {
          itemCols.forEach((colIdx) => {
            nonMatchedList.push({
              colIdx: colIdx,
              type: nonMatch.type,
              number: String(nonMatch.number || "").trim(),
              reason: nonMatch.reason
            });
          });
        });
      }
    });

    // 2. Extract from root nonMatchedBomLine if present
    const rootNonMatched = result.nonMatchedBomLine || result.nonMatchedBomLines || [];
    if (Array.isArray(rootNonMatched)) {
      rootNonMatched.forEach((nonMatch) => {
        flaggedColIndexes.forEach((colIdx) => {
          nonMatchedList.push({
            colIdx: colIdx,
            type: nonMatch.type,
            number: String(nonMatch.number || "").trim(),
            reason: nonMatch.reason
          });
        });
      });
    }

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      const sectionColorMap = {
        "1 - Style Information": "#F2F2F2",
        "2 - Shell Fabric Information": "#F7ECE6",
        "2 - Yarn Information": "#F7ECE6",
        "3 - Other Fabric Information": "#FCE4D6",
        "3 - Other Yarn Information": "#FCE4D6",
        "4 - Sewing/Labeling Trims": "#FCE4D6",
        "4 - Sewing/Labeling Accessories": "#FCE4D6",
        "5 - Sewing/Labeling Accessories": "#FCE4D6",
        "5 - Finishing/Packing Accessories": "#FCE4D6",
        "6 - Finishing/Packing Trims": "#FCE4D6",
        "6 - Sewing/Labeling Trims": "#FCE4D6",
        "7 - Finishing/Packing Accessories": "#FCE4D6",
        "7 - Finishing/Packing Trims": "#FCE4D6",
        "8 - Embellishments": "#FCE4D6",
        "9 - Print & Embroidery Information": "#FCE4D6",
        "10 - Wash Information": "#FCE4D6",
        "11 - BOM Cost": "#F4B084",
        "12 - Operations Cost": "#76C44B",
        "13 - Fixed Cost": "#B3429A",
        "14 - Total Cost": "#9BC2E6",
        "15 - FOB Price": "#9BC2E6",
        "16 - Margin": "#DDEBF7",
        "17 - DDP Price": "#DDEBF7",
        "18 - Action": "#FFFFFF",
      };

      // Reset trim number cell background colors for flagged columns to their section color
      const trimNumberRows = [];
      fieldMap.forEach((field) => {
        if (field._trimGroupCategory !== undefined && field._trimProp === "trimNumber") {
          const rowIdx = findSheetRowIndex(allData, field.section, field.label, field._trimGroupSlot);
          if (rowIdx !== -1) {
            trimNumberRows.push({
              rowIdx: rowIdx,
              section: field.section,
              slot: field._trimGroupSlot
            });
          }
        }
      });

      // Reset material/fabric/yarn number cell background colors for flagged columns to their section color
      const materialNumberRows = [];
      fieldMap.forEach((field) => {
        if (field.key === "mainFabricFabricName") {
          const rowIdx = findSheetRowIndex(allData, field.section, field.label, 0);
          if (rowIdx !== -1) {
            materialNumberRows.push({ rowIdx: rowIdx, section: field.section });
          }
        } else if (field.key && field.key.startsWith("fabricModel[") && field.key.endsWith(".fabricNumber")) {
          const match = field.key.match(/fabricModel\[(\d+)\]/);
          const slotIndex = match ? parseInt(match[1]) : 0;
          const rowIdx = findSheetRowIndex(allData, field.section, field.label, slotIndex);
          if (rowIdx !== -1) {
            materialNumberRows.push({ rowIdx: rowIdx, section: field.section });
          }
        }
      });

      flaggedColIndexes.forEach((colIdx) => {
        // Reset Trim cells to section color
        trimNumberRows.forEach((item) => {
          const cell = sheet.getRangeByIndexes(item.rowIdx, colIdx, 1, 1);
          cell.format.fill.color = sectionColorMap[item.section] || "#FFFFFF";
        });
        // Reset Material cells to section color
        materialNumberRows.forEach((item) => {
          const cell = sheet.getRangeByIndexes(item.rowIdx, colIdx, 1, 1);
          cell.format.fill.color = sectionColorMap[item.section] || "#FFFFFF";
        });
      });

      // Highlight non-matched items in red
      nonMatchedList.forEach((nonMatch) => {
        const colIdx = nonMatch.colIdx;
        const targetNumber = nonMatch.number;
        const type = String(nonMatch.type || "").trim().toLowerCase();

        if (type === "trim") {
          trimNumberRows.forEach((item) => {
            const cellValue = String(allData[item.rowIdx][colIdx] || "").trim();
            if (cellValue === targetNumber) {
              const cell = sheet.getRangeByIndexes(item.rowIdx, colIdx, 1, 1);
              cell.format.fill.color = "#F59393";
              console.log(`Highlighting non-matched trim cell at row ${item.rowIdx}, col ${colIdx} with red.`);
            }
          });
        } else if (type === "material" || type === "fabric" || type === "yarn") {
          materialNumberRows.forEach((item) => {
            const cellValue = String(allData[item.rowIdx][colIdx] || "").trim();
            if (cellValue === targetNumber) {
              const cell = sheet.getRangeByIndexes(item.rowIdx, colIdx, 1, 1);
              cell.format.fill.color = "#F59393";
              console.log(`Highlighting non-matched material cell at row ${item.rowIdx}, col ${colIdx} with red.`);
            }
          });
        }
      });

      flaggedColIndexes.forEach((colIdx) => {
        const status = colStatusMap[colIdx];
        const msgCell = sheet.getRangeByIndexes(actionRowIndex + 1, colIdx, 1, 1);

        if (status.isSuccess) {
          const msg = status.messages.length > 0 ? status.messages.join("\n") : "Updated successfully.";
          msgCell.values = [[msg]];
          console.log(`Costing col ${colIdx} updated successfully: ${msg}`);
        } else if (status.isFailure) {
          const msg = status.messages.length > 0 ? status.messages.join("\n") : "Update failed.";
          msgCell.values = [[msg]];
          console.log(`Costing col ${colIdx} failed to update: ${msg}`);
        }
      });
      await context.sync();
    });

    // showWarning("Upload completed successfully.");
  } catch (error) {
    console.error("UploadData error: ", error);
  } finally {
    hideLoader();
  }
});

function getExpiryTime(seconds) {
  const now = new Date(); // Current time
  const expiry = new Date(now.getTime() + seconds * 1000); // Add seconds
  return expiry; // Returns human-readable format
}
setInterval(() => {
  if (!tokenExpiryDate) return;
  const now = new Date();
  const diffMs = tokenExpiryDate - now;
  const minutesLeft = diffMs / (1000 * 60);
  if (minutesLeft <= 10) {
    // showAuthModal();
  }
}, 30000); // every 30 seconds

// function showAuthModal() {
//   const modal = document.getElementById("authModal");
//   modal.style.display = "block";
//   document.getElementById("authNowBtn").onclick = function () {
//     modal.style.display = "none";
//     document.getElementById("openDialog").click(); // trigger Authenticate
//   };
// }

function showLoader() {
  const loaderEl = document.getElementById("loader");
  if (loaderEl) {
    loaderEl.style.display = "block";
  }
}
function hideLoader() {
  const loaderEl = document.getElementById("loader");
  if (loaderEl) {
    loaderEl.style.display = "none";
  }
}

function showWarning(message) {
  const modalEl = document.getElementById("moduleWarningModal1");
  modalEl.querySelector(".modal-body").textContent = message;
  if (!window.warningModal) {
    window.warningModal = new bootstrap.Modal(modalEl);
  }
  window.warningModal.show();
}
function validateSelections() {
  const selectedSeasonId = document.getElementById("seasonSelect")?.value || "None";
  const selectedDivisionId = document.getElementById("divisionSelect")?.value || "None";
  const selectedBrandId = document.getElementById("brandSelect")?.value || "None";
  const selectedGenderId = document.getElementById("genderSelect")?.value || "None";
  const selectedCategoryId = document.getElementById("categorySelect")?.value || "None";

  let missingFields = [];
  if (!selectedSeasonId || selectedSeasonId === "None" || selectedSeasonId === "") {
    missingFields.push("Season");
  }
  if (!selectedDivisionId || selectedDivisionId === "None" || selectedDivisionId === "") {
    missingFields.push("Department ");
  }
  if (!selectedBrandId || selectedBrandId === "None" || selectedBrandId === "") {
    missingFields.push("Collection Line ");
  }
  if (!selectedGenderId || selectedGenderId === "None" || selectedGenderId === "") {
    missingFields.push("Gender");
  }
  if (!selectedCategoryId || selectedCategoryId === "None" || selectedCategoryId === "") {
    missingFields.push("Category");
  }
  if (missingFields.length > 0) {
    const message = "Please select " + missingFields.join(", ") + ".";
    showWarning(message);
    return false;
  }
  return true;
}

let isButtonToggling = false;

async function toggleButtonVisibility() {
  if (isButtonToggling) return;
  isButtonToggling = true;
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load(["isNullObject", "rowCount", "columnCount"]);
      await context.sync();

      if (usedRange.isNullObject || usedRange.rowCount < 2) {
        console.log("toggleButtonVisibility: usedRange is null or rowCount < 2");
        document.getElementById("UploadData").style.display = "none";
        return;
      }

      const allDataRange = sheet.getRangeByIndexes(0, 0, usedRange.rowCount, usedRange.columnCount);
      allDataRange.load("values");
      await context.sync();
      const allData = allDataRange.values;
      const totalCols = usedRange.columnCount;

      console.log(`toggleButtonVisibility: totalRows=${usedRange.rowCount}, totalCols=${totalCols}`);

      const actionRowIndex = allData.findIndex((row, idx) => {
        const valA = String(row[0]).trim().toLowerCase();
        const valB = String(row[1]).trim().toLowerCase();
        const match = valA === "action for update" || valA === "update y/n" || valA === "18 - action" ||
          valB === "action for update" || valB === "update y/n" || valB === "18 - action" ||
          valA.includes("action") || valA.includes("update") ||
          valB.includes("action") || valB.includes("update");
        if (match) {
          console.log(`toggleButtonVisibility: found action row at index ${idx}. ColA='${row[0]}', ColB='${row[1]}'`);
        }
        return match;
      });

      if (actionRowIndex === -1) {
        console.log("toggleButtonVisibility: Action row not found in sheet.");
        document.getElementById("UploadData").style.display = "none";
        return;
      }

      const actionRowValues = allData[actionRowIndex];
      const flaggedColIndexes = [];

      for (let colIdx = 2; colIdx < totalCols; colIdx++) {
        const val = actionRowValues[colIdx];
        // console.log(`toggleButtonVisibility: col ${colIdx} value is '${val}'`);
        const sVal = String(val || "").trim().toUpperCase();
        if (val === 1 || val === "1" || sVal === "Y" || sVal === "y") {
          flaggedColIndexes.push(colIdx);
        }
      }
      const hasFlag = flaggedColIndexes.length > 0;
      document.getElementById("UploadData").style.display = hasFlag ? "inline-block" : "none";
      console.log("Upload button visibility updated:", hasFlag ? "VISIBLE" : "HIDDEN", "Flagged columns:", flaggedColIndexes);
    });
  } catch (error) {
    if (error.message && error.message.indexOf("cell-editing mode") !== -1) {
      console.warn("toggleButtonVisibility: Excel is in cell-editing mode. Operation deferred.");
    } else {
      console.error("toggleButtonVisibility error:", error);
    }
  } finally {
    isButtonToggling = false;
  }
}

async function ensureStyleResponseData(allData, totalCols) {
  let debugLog = [];
  debugLog.push("ensureStyleResponseData started.");

  if (StyleResponseData && StyleResponseData.length > 0) {
    debugLog.push("StyleResponseData already populated.");
    lastDebugLog = debugLog.join(" | ");
    return true;
  }

  // 1. Find Style Number row in allData
  const styleNumberRowIndex = allData.findIndex((row) => {
    const lblA = String(row[0] || "").trim().toLowerCase();
    const lblB = String(row[1] || "").trim().toLowerCase();
    return lblA === "style number" || lblB === "style number" || lblA.includes("style number") || lblB.includes("style number");
  });
  debugLog.push("styleNumberRowIndex: " + styleNumberRowIndex);
  if (styleNumberRowIndex === -1) {
    const sampleLabels = allData.slice(0, 10).map((row, idx) => `Row ${idx}: ColA='${row[0]}', ColB='${row[1]}'`);
    debugLog.push("First 10 rows: " + sampleLabels.join("; "));
    lastDebugLog = debugLog.join(" | ");
    return false;
  }

  // 2. Read Style Numbers from sheet columns
  const styleNumbersInSheet = [];
  for (let colIdx = 2; colIdx < totalCols; colIdx++) {
    const styleNum = String(allData[styleNumberRowIndex][colIdx] || "").trim();
    if (styleNum && !styleNumbersInSheet.includes(styleNum)) {
      styleNumbersInSheet.push(styleNum);
    }
  }
  debugLog.push("styleNumbersInSheet: " + JSON.stringify(styleNumbersInSheet));
  if (styleNumbersInSheet.length === 0) {
    lastDebugLog = debugLog.join(" | ");
    return false;
  }
  // 3. Get dropdown selections
  const selectedSeasonId = document.getElementById("seasonSelect")?.value || "None";
  const selectedDivisionId = document.getElementById("divisionSelect")?.value || "None";
  const selectedBrandId = document.getElementById("brandSelect")?.value || "None";
  const selectedGenderId = document.getElementById("genderSelect")?.value || "None";
  const selectedCategoryId = document.getElementById("categorySelect")?.value || "None";

  debugLog.push(`Dropdowns - Season: ${selectedSeasonId}, Div: ${selectedDivisionId}, Brand: ${selectedBrandId}, Gender: ${selectedGenderId}, Category: ${selectedCategoryId}`);

  if (
    !selectedSeasonId || selectedSeasonId === "None" || selectedSeasonId === "" ||
    !selectedDivisionId || selectedDivisionId === "None" || selectedDivisionId === "" ||
    !selectedBrandId || selectedBrandId === "None" || selectedBrandId === "" ||
    !selectedGenderId || selectedGenderId === "None" || selectedGenderId === "" ||
    !selectedCategoryId || selectedCategoryId === "None" || selectedCategoryId === ""
  ) {
    debugLog.push("Dropdowns not fully selected.");
    lastDebugLog = debugLog.join(" | ");
    return false;
  }

  // 4. Fetch style list from server to map styleNumber -> styleId
  try {
    const accessToken = document.getElementById("access_token").innerText;
    const styleListUrl = apiURL + `/api/StyleLists/GetStyles?seasonId=${selectedSeasonId}&brandId=${selectedBrandId}&divisionId=${selectedDivisionId}&genderId=${selectedGenderId}&categoryId=${selectedCategoryId}`;
    debugLog.push("Fetching style list from: " + styleListUrl);
    const response = await fetch(styleListUrl, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json",
      },
    });
    debugLog.push("Response status: " + response.status);
    if (!response.ok) {
      lastDebugLog = debugLog.join(" | ");
      return false;
    }
    const stylesList = await response.json();
    debugLog.push("stylesList count: " + stylesList.length);
    if (stylesList.length > 0) {
      const sampleStyles = stylesList.slice(0, 3).map(s => `StyleId=${s.StyleId}, Code=${s.StyleCode}, Num=${s.StyleNumber || s.styleNumber}`);
      debugLog.push("Sample styles: " + sampleStyles.join("; "));
    }

    // 5. Match Style Numbers in sheet to get corresponding Style IDs
    const styleIdsToFetch = [];
    styleNumbersInSheet.forEach((styleNum) => {
      const found = stylesList.find(
        (s) => String(s.StyleCode || s.styleCode || s.StyleNumber || s.styleNumber || "").trim().toLowerCase() === styleNum.toLowerCase()
      );
      if (found && (found.StyleId || found.styleId)) {
        styleIdsToFetch.push(found.StyleId || found.styleId);
      }
    });
    debugLog.push("styleIdsToFetch: " + JSON.stringify(styleIdsToFetch));

    if (styleIdsToFetch.length === 0) {
      lastDebugLog = debugLog.join(" | ");
      return false;
    }

    // 6. Fetch details for each Style ID in parallel
    const fetchPromises = styleIdsToFetch.map((id) => {
      const styleDetailsUrl = apiURL + `/api/Styles/StyleDetails?styleId=${id}`;
      return fetch(styleDetailsUrl, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json",
        },
      }).then((res) => {
        if (!res.ok) throw new Error("Failed to fetch style details");
        return res.json();
      });
    });

    const results = await Promise.all(fetchPromises);
    let combinedStyleDetails = [];
    results.forEach((data) => {
      let styleDetails = data.styleData || data.StyleData || [];
      combinedStyleDetails = combinedStyleDetails.concat(styleDetails);
    });
    debugLog.push("combinedStyleDetails count: " + combinedStyleDetails.length);

    combinedStyleDetails.forEach((costing) => {
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

    if (combinedStyleDetails.length > 0) {
      StyleResponseData = combinedStyleDetails;
      try {
        localStorage.setItem("StyleResponseData", JSON.stringify(combinedStyleDetails));
      } catch (e) {
        console.error("Failed to save StyleResponseData to localStorage:", e);
      }
      if (typeof Office !== "undefined" && Office.context?.document?.settings) {
        try {
          Office.context.document.settings.set("StyleResponseData", combinedStyleDetails);
          Office.context.document.settings.saveAsync();
        } catch (e) {
          console.error("Failed to save StyleResponseData to Office settings:", e);
        }
      }
      debugLog.push("StyleResponseData loaded successfully.");
      lastDebugLog = debugLog.join(" | ");
      return true;
    }
  } catch (e) {
    debugLog.push("Exception: " + e.message);
    console.error("Error in ensureStyleResponseData:", e);
  }
  lastDebugLog = debugLog.join(" | ");
  return false;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    buildFieldMap,
    mapToCostingUpdateRequest
  };
}