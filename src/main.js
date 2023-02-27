const input = document.querySelector("input");
const btn = document.querySelector("button");
const container = document.querySelector("#container");
const download_btn = document.querySelector("a");

btn.addEventListener("click", () => {
    const workbook = XLSX.readFile(input.files[0].path);
    const MPD_sheet = workbook.Sheets["MPD"];

    const arr = XLSX.utils.sheet_to_json(MPD_sheet,
        {header: [
            "REV CODE",
            "SECTION",
            "TASK NUMBER",
            "SOURCE TASK REFERENCE",
            "ACCESS",
            "PREPARATION",
            "ZONE",
            "DESCRIPTION",
            "SKILL CODE",
            "TASK CODE",
            "SAMPLE THRESHOLD",
            "SAMPLE INTERVAL",
            "100% THRESHOLD",
            "100% INTERVAL",
            "SOURCE",
            "TCI",
            "RSC",
            "TPS",
            "REFERENCE",
            "MEN",
            "TASKM.H.",
            "ACCESSM.H.",
            "PREP.M.H.",
            "APPLICABILITY"
        ]
    });

    
    container.textContent = arr[2]["SECTION"];

    const download_worksheet = XLSX.utils.json_to_sheet(arr.slice(2));
    const download_workbook = XLSX.utils.book_new();

    // XLSX.utils.book_append_sheet(download_workbook, download_worksheet, "First");
    download_workbook.SheetNames.push("First");
    download_workbook.Sheets["First"] = download_worksheet;

    let name = "file";
    XLSX.writeFile(download_workbook, `${name}.xlsx`);
    download_btn.href = `../${name}.xlsx`;
    download_btn.classList.remove("btn-secondary");
    download_btn.classList.add("btn-success");
    download_btn.classList.remove("disabled");
    
    // console.log(download_workbook);
});