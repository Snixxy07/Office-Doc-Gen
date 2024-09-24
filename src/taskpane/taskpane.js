/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// Default FOP data
const defaultFop = {
  fop: "Проценко Юрій Ігорович",
  sex: "m",
  inn: "3289205817",
  registrationDate: "06.12.2023",
  registrationNumber: "2005560000000181053",
  address: "65025, Одеська обл., місто Одеса, пр. Добровольського, будинок 137, квартира 55",
  accountNumber: "UA113071230000026004011398566",
  bank: "БАНК ВОСТОК",
  bankAbbreviation: "ПАТ",
};

const partsToReplace = [
  { searchText: "Додаток або із залученням уповноважених ними Кур’єрів", replacementText: "Додаток" },
  {
    searchText: "зареєстрована у Додатку  особа, яка самостійно та на власний ризик займається доставкою",
    replacementText: "особа, яка за дорученням Принципала здійснює доставку",
  },
  {
    searchText:
      "Товарів між Принципалом та Користувачем, прийняття оплати за Товари, замовлені Користувачами у Принципалів, організації доставки",
    replacementText:
      "Товарів між Принципалом та Користувачем з доставкою Товарів Кур'єрами, прийняття оплати за Товари, замовлені Користувачами у Принципалів, та оплати за доставку Кур'єрами",
  },
  {
    searchText:
      "які Принципал пропонує Користувачам через Додаток. Для отримання зазначених платежів Агент використовує мережу Інтернет та укладає договори з фінансовими компаніями",
    replacementText:
      "які Принципал пропонує Користувачам через Додаток, та за послуги з доставки таких Товарів Кур'єрами. Для отримання зазначених платежів Агент використовує мережу Інтернет та укладає договори з фінансовими установами",
  },
];

const appId = "-fafvhitsjq-uc.a.run.app";

// Initialize the Office Add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initializeEventListeners();
    populateOurFopSelect().catch((error) => console.error("Error populating FOP select:", error));
  }
});

// Initialize all event listeners
function initializeEventListeners() {
  document.querySelector("#addFopBtn").onclick = () => tryCatch(showAddFopForm);
  document.querySelector("#closeAddFopForm").onclick = () => tryCatch(closeAddFopForm);

  const fopInput = document.getElementById("fop");
  fopInput.addEventListener("blur", formatFopName);

  const contractDateInput = document.getElementById("contractDate");
  contractDateInput.addEventListener("change", updateContractEndDate);

  document.getElementById("replaceForm").onsubmit = (event) => {
    event.preventDefault();
    console.log("Form submitted");
    tryCatch(replaceData);
  };

  document.getElementById("addFopForm").onsubmit = (event) => {
    event.preventDefault();
    tryCatch(saveFormData);
  };

  const contractNumberInput = document.getElementById("contractNumber");
  loadLastContractNumber().then((lastNumber) => {
    contractNumberInput.value = lastNumber;
  });
  contractNumberInput.onblur = (event) => {
    handleContractNumberChange(event);
  };

  const ourFopSelect = document.getElementById("ourFop");
  const fopToReplaceSelect = document.getElementById("replaceFop");

  ourFopSelect.addEventListener("change", (event) => {
    localStorage.setItem("lastSelectedOurFop", event.target.value);
  });

  fopToReplaceSelect.addEventListener("change", (event) => {
    localStorage.setItem("lastSelectedReplaceFop", event.target.value);
  });
}

async function loadLastContractNumber() {
  try {
    const localLastNum = localStorage.getItem("lastContractNumber") || "";
    const response = await fetch(`https://getcontractnumber${appId}`);
    const data = await response.json();

    if (data.lastNumber) {
      localStorage.setItem("lastContractNumber", data.lastNumber);
      return data.lastNumber;
    }

    return localLastNum;
  } catch (error) {
    console.error("Error fetching last contract number:", error);
    return localStorage.getItem("lastContractNumber") || "";
  }
}

function handleContractNumberChange(event) {
  saveContractNumber(event.target.value);
}

function saveContractNumber(value) {
  fetch(`https://updatecontractnumber${appId}?number=${value}`).then((response) => console.log(response.status));
  localStorage.setItem("lastContractNumber", value);
}

// UI Functions
function showAddFopForm() {
  document.getElementById("addFopForm").classList.remove("hidden");
}

function closeAddFopForm() {
  document.getElementById("addFopForm").classList.add("hidden");
}

// Form Handling Functions
function formatFopName() {
  this.value = this.value
    .toLowerCase()
    .split(" ")
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}

function updateContractEndDate() {
  const contractDateInput = document.getElementById("contractDate");
  const contractEndDateInput = document.getElementById("contractEndDate");

  if (contractDateInput.value) {
    const startDate = new Date(contractDateInput.value);
    const endDate = new Date(startDate.getFullYear() + 1, startDate.getMonth(), startDate.getDate() + 1);
    contractEndDateInput.value = endDate.toISOString().split("T")[0];
  } else {
    contractEndDateInput.value = "";
  }
}

function validateFormData(formData) {
  for (const [key, value] of Object.entries(formData)) {
    if (!value) return false;
  }

  if (!/^\d{10}$/.test(formData.inn)) return false;
  if (!/^UA\d{27}$/.test(formData.accountNumber)) return false;

  return true;
}

/* function saveFormData() {
  const formData = {
    fop: document.getElementById("fop").value,
    sex: document.getElementById("sex").value,
    inn: document.getElementById("inn").value,
    registrationDate: document.getElementById("registrationDate").value,
    registrationNumber: document.getElementById("registrationNumber").value,
    address: document.getElementById("address").value,
    accountNumber: document.getElementById("accountNumber").value,
    bank: document.getElementById("bank").value,
    bankAbbreviation: document.getElementById("bankAbbreviation").value,
  };

  if (validateFormData(formData)) {
    let fopDataArray = JSON.parse(localStorage.getItem("fopDataArray")) || [];
    const existingIndex = fopDataArray.findIndex((item) => item.inn === formData.inn);

    if (existingIndex !== -1) {
      fopDataArray[existingIndex] = formData;
    } else {
      fopDataArray.push(formData);
    }

    localStorage.setItem("fopDataArray", JSON.stringify(fopDataArray));
    console.log("Settings saved.");
    document.getElementById("addFopForm").reset();
    closeAddFopForm();
    populateOurFopSelect();
  }
} */

async function saveFormData() {
  const formData = {
    fop: document.getElementById("fop").value,
    sex: document.getElementById("sex").value,
    inn: document.getElementById("inn").value,
    registrationDate: document.getElementById("registrationDate").value,
    registrationNumber: document.getElementById("registrationNumber").value,
    address: document.getElementById("address").value,
    accountNumber: document.getElementById("accountNumber").value,
    bank: document.getElementById("bank").value,
    bankAbbreviation: document.getElementById("bankAbbreviation").value,
  };

  if (validateFormData(formData)) {
    try {
      const response = await fetch(`https://storeFopData${appId}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(formData),
      });

      const result = await response.json();

      if (response.ok) {
        console.log("FOP data saved to server:", result.message);
        document.getElementById("addFopForm").reset();
        closeAddFopForm();
        populateOurFopSelect();
      } else if (response.status === 409) {
        console.warn("FOP with this INN already exists:", result.message);
        // Handle the case where the FOP already exists (e.g., show a message to the user)
      } else {
        console.error("Failed to save FOP data to server:", result.message);
      }
    } catch (error) {
      console.error("Error saving FOP data:", error);
    }
  }
}

// FOP Data Handling Functions
/* function getAllFops() {
  const fopDataArray = JSON.parse(localStorage.getItem("fopDataArray")) || [];
  return fopDataArray.reduce((acc, fop) => {
    acc[fop.inn] = fop;
    return acc;
  }, {});
} */

async function getAllFops() {
  try {
    const response = await fetch(`https://getAllFops${appId}`);
    const data = await response.json();

    if (response.ok) {
      const fops = {};
      data.fops.forEach((fop) => {
        fops[fop.inn] = fop;
      });
      return fops;
    } else {
      console.error("Failed to fetch FOPs:", data.message);
      return {};
    }
  } catch (error) {
    console.error("Error fetching FOPs:", error);
    return {};
  }
}

async function populateOurFopSelect() {
  const fops = await getAllFops();
  const ourFopSelect = document.getElementById("ourFop");
  const fopToReplaceSelect = document.getElementById("replaceFop");
  ourFopSelect.innerHTML = "";
  fopToReplaceSelect.innerHTML = "";

  // Get last selected values from local storage
  const lastSelectedOurFop = localStorage.getItem("lastSelectedOurFop");
  const lastSelectedReplaceFop = localStorage.getItem("lastSelectedReplaceFop");

  let isFirstOption = true;
  for (const inn in fops) {
    const ourFopOption = document.createElement("option");
    const fopToReplaceOption = document.createElement("option");

    ourFopOption.value = inn;
    ourFopOption.textContent = fops[inn].fop;
    fopToReplaceOption.value = inn;
    fopToReplaceOption.textContent = fops[inn].fop;

    // Set selected option based on last selected values
    if (inn === lastSelectedOurFop) {
      ourFopOption.selected = true;
      isFirstOption = false;
    }
    if (inn === lastSelectedReplaceFop) {
      fopToReplaceOption.selected = true;
      isFirstOption = false;
    }

    // If no saved selection, select the first option
    if (isFirstOption) {
      ourFopOption.selected = true;
      fopToReplaceOption.selected = true;
      isFirstOption = false;
    }

    ourFopSelect.appendChild(ourFopOption);
    fopToReplaceSelect.appendChild(fopToReplaceOption);
  }

  // Add event listeners to save selected values
  ourFopSelect.addEventListener("change", (event) => {
    localStorage.setItem("lastSelectedOurFop", event.target.value);
  });

  fopToReplaceSelect.addEventListener("change", (event) => {
    localStorage.setItem("lastSelectedReplaceFop", event.target.value);
  });

  if (ourFopSelect.options.length === 0) {
    const defaultOption1 = document.createElement("option");
    const defaultOption2 = document.createElement("option");
    defaultOption1.value = "";
    defaultOption1.textContent = "Наш ФОП";
    defaultOption2.value = "";
    defaultOption2.textContent = "Наш ФОП";
    ourFopSelect.appendChild(defaultOption1);
    fopToReplaceSelect.appendChild(defaultOption2);
  }
}

// Data Replacement Functions
async function replaceData() {
  const selectedFopInn = document.getElementById("ourFop").value;
  const fopToReplaceInn = document.getElementById("replaceFop").value;
  const contractNumber = document.getElementById("contractNumber").value;
  const contractDate = document.getElementById("contractDate").value;
  const contractEndDate = document.getElementById("contractEndDate").value;

  const fops = await getAllFops();
  const selectedFop = fops[selectedFopInn];
  const fopToReplace = fops[fopToReplaceInn];

  if (!selectedFop || !fopToReplace) {
    console.error("Selected FOP or FOP to replace not found");
    return;
  }

  await replaceFopData(selectedFop, fopToReplace);
  await replaceContractData(contractNumber, contractDate, contractEndDate);
  await replaceNeededParts(partsToReplace);

  console.log("Data replacement completed.");
}

async function replaceFopData(selectedFop, fopToReplace) {
  await replaceText(fopToReplace.fop, selectedFop.fop);
  await replaceText(fopToReplace.inn, selectedFop.inn);
  await replaceText(proceedSex(fopToReplace.sex), proceedSex(selectedFop.sex), true);
  await replaceText(fopToReplace.registrationDate, selectedFop.registrationDate);
  await replaceText(fopToReplace.registrationNumber, selectedFop.registrationNumber);
  await replaceText(fopToReplace.address, selectedFop.address);
  await replaceText(fopToReplace.accountNumber, selectedFop.accountNumber);
  await replaceText(
    formatBankName(fopToReplace.bank, fopToReplace.bankAbbreviation),
    formatBankName(selectedFop.bank, selectedFop.bankAbbreviation)
  );
}

async function replaceContractData(contractNumber, contractDate, contractEndDate) {
  if (contractNumber) {
    await replaceTextRegex("([0-9]@)/24", contractNumber + "/24");
  }
  if (contractDate) {
    await replaceTextRegex("«[0-9]{2}» ([!0-9]@) 2024", formatDateUkrainian(contractDate));
  }
  if (contractEndDate) {
    await replaceTextRegex("«[0-9]{2}» ([!0-9]@) 2025", formatDateUkrainian(contractEndDate));
  }
}

// Utility Functions
function proceedSex(sex) {
  return sex === "m" ? "який" : "яка";
}

function formatBankName(bankName, bankAbbreviation) {
  return `${bankAbbreviation} «${bankName}»`;
}

function formatDateUkrainian(dateString) {
  const date = new Date(dateString);
  const day = date.getDate().toString().padStart(2, "0");
  const year = date.getFullYear();
  const months = [
    "січня",
    "лютого",
    "березня",
    "квітня",
    "травня",
    "червня",
    "липня",
    "серпня",
    "вересня",
    "жовтня",
    "листопада",
    "грудня",
  ];
  return `«${day}» ${months[date.getMonth()]} ${year}`;
}

// Word Document Manipulation Functions
async function replaceText(searchText, replacementText, replaceFirstOnly = false) {
  await Word.run(async (context) => {
    try {
      const doc = context.document.body;
      const searchResults = doc.search(searchText, { matchCase: false, matchWholeWord: true });
      searchResults.load("items");
      await context.sync();

      if (searchResults.items.length > 0) {
        if (replaceFirstOnly) {
          searchResults.items[0].insertText(replacementText, Word.InsertLocation.replace);
        } else {
          searchResults.items.forEach((result) => {
            result.insertText(replacementText, Word.InsertLocation.replace);
          });
        }
        console.log(
          `Replaced ${replaceFirstOnly ? "first occurrence" : searchResults.items.length + " occurrences"} of "${searchText}" with "${replacementText}"`
        );
      } else {
        console.log(`No matches found for "${searchText}"`);
      }

      await context.sync();
    } catch (error) {
      console.log("Error in replaceText:" + error);
    }
  });
}

async function replaceTextRegex(searchPattern, replacementText, replaceFirstOnly = false) {
  await Word.run(async (context) => {
    try {
      const doc = context.document.body;
      const searchResults = doc.search(searchPattern, { matchWildcards: true });
      searchResults.load("items");
      await context.sync();

      if (searchResults.items.length > 0) {
        if (replaceFirstOnly) {
          searchResults.items[0].insertText(replacementText, Word.InsertLocation.replace);
        } else {
          searchResults.items.forEach((result) => {
            result.insertText(replacementText, Word.InsertLocation.replace);
          });
        }
        console.log(
          `Replaced ${replaceFirstOnly ? "first occurrence" : searchResults.items.length + " occurrences"} matching "${searchPattern}" with "${replacementText}"`
        );
      } else {
        console.log(`No matches found for "${searchPattern}"`);
      }

      await context.sync();
    } catch (error) {
      console.log("Error in replaceTextRegex: " + error);
    }
  });
}

// Error Handling
async function tryCatch(callback) {
  try {
    await callback();
    console.log("Action completed.");
  } catch (error) {
    console.error(error);
  }
}

async function replaceNeededParts(parts) {
  await tryCatch(async () => {
    for (const part of parts) {
      await replaceText(part.searchText, part.replacementText, true);
    }
  });
}
