Office.onReady(function () {
  $("#placeholderInput").on("input", showSuggestions);
  $("#insertPlaceholder").on("click", insertPlaceholder);
  $("#startMailMerge").on("click", performMailMerge);
});
let debounceTimeout;
var selectedFields = []; // Store selected item globally

async function showSuggestions() {
  clearTimeout(debounceTimeout);

  debounceTimeout = setTimeout(async () => {
    let input = $("#placeholderInput").val().trim();
    let token = $("#apiToken").val().trim();
    let suggestionsList = $("#suggestions");

    suggestionsList.empty().hide();

    if (input.length < 2) return;

    try {
      const response = await fetch(
        `https://localhost:8081/api/Suggestion/Query?year=2024&query=${encodeURIComponent(input)}`,
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
            console.log(ph);
            selectedFields.push(ph);
            $("#placeholderInput").val(ph.name); // ✅ insert name into input
            suggestionsList.hide();
          });

        suggestionsList.append(item);
      });

      if (placeholders.length > 0) {
        suggestionsList.show();
      }
    } catch (error) {
      console.error("Error fetching suggestions:", error);
      $("#notification-body").html("An error occurred while fetching suggestions");
    }
  }, 800); // Debounce delay
}

async function insertPlaceholder() {
  let input = $("#placeholderInput").val();
  if (selectedFields.some((x) => x.name == input)) {
    const temp = selectedFields.filter((x) => x.name == input);

    await Word.run(async (context) => {
      let selection = context.document.getSelection();
      selection.insertText("{{" + temp[0].encodeName + "}}", Word.InsertLocation.replace);
      await context.sync();
    });
  }
}

async function performMailMerge() {
  let input = $("#placeholderInput").val();
  if (selectedFields.some((x) => x.name == input)) {
    const temp = selectedFields.filter((x) => x.name == input);

    await Word.run(async (context) => {
      let selection = context.document.getSelection();
      selection.insertText(temp[0].value, Word.InsertLocation.replace);
      await context.sync();
    });
  }
}
