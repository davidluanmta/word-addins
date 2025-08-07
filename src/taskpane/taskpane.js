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
// const webServer = "https://localhost:8081";
const webServer = "https://report-api.ueh.edu.vn/";
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
      const response = await fetch(
        `${webServer}/api/Suggestion/Query?year=${year}&query=${encodeURIComponent(input)}`,
        {
          method: "GET",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
        }
      );

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

  const response = await fetch(`${webServer}/api/Suggestion/FetchTableInfo?id=${id}`, {
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
  let input = $("#placeholderInput").val();
  if (selectedFields.some((x) => x.name == input)) {
    const temp = selectedFields.filter((x) => x.name == input);
    console.log(temp);
    await Word.run(async (context) => {
      let selection = context.document.getSelection();
      selection.insertText(temp[0].value, Word.InsertLocation.replace);
      await context.sync();
    });
  }
}
