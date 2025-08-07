Office.onReady(function () {
  $("#placeholderInput").on("input", showSuggestions);
  $("#insertPlaceholder").on("click", insertPlaceholder);
  $("#startMailMerge").on("click", performMailMerge);
  $(document).ready(function () {
    $("#columns").select2({
      placeholder: "Chọn các cột để hiển thị",
      allowClear: true,
      width: "resolve",
    });
  });
});
let debounceTimeout;
var selectedFields = []; // Store selected item globally
var selectedColumns = []; // Store selected item globally

let isTable = false;
// const webServer = "https://localhost:8081/api/Suggestion";
const webServer = "https://report-api.ueh.edu.vn/api/Suggestion";
let token = ""; // Set this from user input if needed
async function showSuggestions() {
  clearTimeout(debounceTimeout);

  debounceTimeout = setTimeout(async () => {
    let year = $("#selectedYear").val().trim();
    let input = $("#placeholderInput").val().trim();
    token = $("#apiToken").val().trim();
    let suggestionsList = $("#suggestions");

    suggestionsList.empty().hide();

    if (input.length < 2) return;

    try {
      const response = await fetch(`${webServer}/Query?year=${year}&query=${encodeURIComponent(input)}`, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
      });

      if (!response.ok) {
        $("#notification-body").html("Failed to fetch suggestions");
        return;
      }

      const { data: placeholders } = await response.json(); // ✅ extract `data` field

      placeholders.forEach((ph) => {
        const item = $("<li>")
          .text(ph.name) // ✅ show name in list
          .css({ cursor: "pointer", padding: "5px" })
          .on("click", function () {
            selectedFields.push(ph);
            $("#placeholderInput").val(ph.name); // ✅ insert name into input
            suggestionsList.hide();
            if (ph.isTable) {
              isTable = ph.isTable;
              $(".table-columns").show();
              fetchTableColumn(ph.id);
            } else {
              $(".table-columns").hide();
              selectedColumns = [];
              isTable = false;
            }
          });

        suggestionsList.append(item);
      });

      if (placeholders.length > 0) {
        suggestionsList.show();
      }
    } catch (error) {
      console.error("Error fetching suggestions:", error);
      $("#notification-body").html("An error occurred while fetching suggestions" + error);
    }
  }, 500); // Debounce delay
}

async function fetchTableColumn(id) {
  const token = $("#apiToken").val().trim();

  const response = await fetch(`${webServer}/FetchTableInfo?id=${id}`, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    $("#notification-body").html("Failed to fetch suggestions");
    return;
  }

  const columnsSelect = $("#columns");
  columnsSelect.empty(); // clear old options
  const { data } = await response.json();

  // Populate Select2 options
  data.forEach((col) => {
    const option = $("<option>").val(col).text(col); // Assuming col is string
    columnsSelect.append(option);
  });

  // Initialize or refresh Select2
  columnsSelect.select2({ width: "100%", placeholder: "Chọn cột dữ liệu" });
  columnsSelect.trigger("change");

  // Handle selection
  columnsSelect.off("change").on("change", function () {
    selectedColumns = $(this).val(); // This is an array of strings
    // Display as comma-separated or styled list
    $("#selectedColumns").text(JSON.stringify(selectedColumns));
  });
}

async function getSelectedFromApi(ids) {
  const response = await fetch(webServer + "/FetchSelected", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ids }),
  });

  const json = await response.json();
  if (json.code === 200 && json.data) {
    return json.data;
  }

  return {};
}
async function fillAllPlaceholdersBatch(callApi) {
  await Word.run(async (context) => {
    const body = context.document.body;

    $("#notification-body").html("Đang xử lý dữ liệu, vui lòng chờ...");

    body.load("text");
    await context.sync();

    const fullText = body.text;
    const regex = /\{\{(\d+)_([\w.]+)(?:\s*\[(.*?)\])?\}\}/g;
    const matches = Array.from(fullText.matchAll(regex));

    const placeholders = [];
    const idsSet = new Set();

    for (const match of matches) {
      const full = match[0];
      const id = match[1];
      const rawCols = match[3];
      const columns = rawCols ? rawCols.split(",").map((c) => c.trim().replace(/^"|"$/g, "")) : null;

      placeholders.push({ full, id, columns });
      idsSet.add(id);
    }

    const ids = Array.from(idsSet);
    const dataMap = await callApi(ids);

    for (const { full, id, columns } of placeholders) {
      const value = dataMap[id];
      if (!value) continue;

      const results = body.search(full, { matchCase: false, matchWholeWord: false });
      context.load(results, "items");
      await context.sync();

      if (results.items.length === 0) continue;
      const range = results.items[0];

      if (Array.isArray(value)) {
        if (value.length === 0) continue;

        // Determine columns
        const cols = columns && columns.length > 0 ? columns : Object.keys(value[0] || {});
        if (cols.length === 0) continue;

        const tableValues = [cols];

        for (const row of value) {
          const rowValues = cols.map((col) => (row[col] !== undefined && row[col] !== null ? row[col].toString() : ""));
          tableValues.push(rowValues);
        }

        const rowCount = tableValues.length;
        const colCount = cols.length;

        // Ensure every row has correct column length
        const allRowsValid = tableValues.every((row) => row.length === colCount);
        if (!allRowsValid) {
          console.error("❌ Dữ liệu không đồng đều số cột:", tableValues);
          continue;
        }

        try {
          range.insertTable(rowCount, colCount, Word.InsertLocation.replace, tableValues);
          await context.sync();
        } catch (e) {
          console.error("❌ Lỗi khi insert table:", full, e, {
            rowCount,
            colCount,
            tableValues,
          });
        }
      } else {
        try {
          range.insertText(value.toString(), Word.InsertLocation.replace);
          await context.sync();
        } catch (e) {
          console.error("❌ Lỗi khi insert text:", full, e);
        }
      }
    }

    $("#notification-body").html("");
  });
}

async function insertPlaceholder() {
  let input = $("#placeholderInput").val();
  if (selectedFields.some((x) => x.name == input)) {
    const temp = selectedFields.filter((x) => x.name == input);
    console.log(temp);

    await Word.run(async (context) => {
      let selection = context.document.getSelection();
      if (isTable) {
        selection.insertText(
          "{{" + temp[0].encodeName + JSON.stringify(selectedColumns) + "}}",
          Word.InsertLocation.replace
        );
      } else {
        selection.insertText("{{" + temp[0].encodeName + "}}", Word.InsertLocation.replace);
      }

      await context.sync();
    });
  }
}

async function performMailMerge() {
  await fillAllPlaceholdersBatch(getSelectedFromApi);
}
