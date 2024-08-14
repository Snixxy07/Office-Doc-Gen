/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.querySelector("#addFopBtn").onclick = () => tryCatch(showAddFopForm);
    document.querySelector("#closeAddFopForm").onclick = () => tryCatch(closeAddFopForm);

    const fopInput = document.getElementById("fop");
    fopInput.addEventListener("blur", formatFopName);

    const contractDateInput = document.getElementById("contractDate");
    contractDateInput.addEventListener("change", updateContractEndDate);

    document.getElementById("addFopForm").addEventListener("submit", function (event) {
      event.preventDefault();
      tryCatch(saveFormData);
    });

    populateOurFopSelect();

    /* const contractNumberInput = document.getElementById("contractNumber");
    contractNumberInput.value = loadLastContractNumber();
    contractNumberInput.addEventListener("input", handleContractNumberChange); */

    //document.querySelector('input[type="submit"]').onclick = () => tryCatch(replaceData);
  }
});

function showAddFopForm() {
  document.getElementById("addFopForm").classList.remove("hidden");
}

function closeAddFopForm() {
  document.getElementById("addFopForm").classList.add("hidden");
}

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
    if (!value) {
      alert(`Пожалуйста, заполните поле ${key}.`);
      return false;
    }
  }

  if (!/^\d{10}$/.test(formData.inn)) {
    alert("ИНН должен содержать 10 цифр.");
    return false;
  }

  if (!/^UA\d{27}$/.test(formData.accountNumber)) {
    alert("Номер счета должен быть в формате UA + 27 цифр, например: UA093071230000026009010584020");
    return false;
  }
  return true;
}

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

function saveFormData() {
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

    // Check if an entry with the same INN already exists
    const existingIndex = fopDataArray.findIndex((item) => item.inn === formData.inn);

    if (existingIndex !== -1) {
      // Update existing entry
      fopDataArray[existingIndex] = formData;
    } else {
      // Add new entry
      fopDataArray.push(formData);
    }

    // Save the updated array
    localStorage.setItem("fopDataArray", JSON.stringify(fopDataArray));
    console.log("Settings saved.");
    document.getElementById("addFopForm").reset();
    closeAddFopForm();
    populateOurFopSelect();
  }
}

function getAllFops() {
  const fopDataArray = JSON.parse(localStorage.getItem("fopDataArray")) || [];
  return fopDataArray.reduce((acc, fop) => {
    acc[fop.inn] = fop;
    return acc;
  }, {});
}

function populateOurFopSelect() {
  const fops = getAllFops();
  const select = document.getElementById("ourFop");
  select.innerHTML = '<option value="">Наш ФОП</option>';

  for (const inn in fops) {
    const option = document.createElement("option");
    option.value = inn;
    option.textContent = fops[inn].fop;
    select.appendChild(option);
  }
}

async function replaceData() {
  await Word.run(async (context) => {
    try {
      const doc = context.document.body;

      const searchResults = doc.search("Проценко Юрій Ігорович", { matchCase: false, matchWholeWord: true });
      searchResults.load("items");
      await context.sync();
      console.log("Found count: " + searchResults.items.length);
      if (searchResults.items.length > 0) {
        searchResults.items.forEach((result) => {
          console.log(`Found '${result.text}'`);
          result.insertText("Юрко", Word.InsertLocation.replace);
        });
      } else {
        console.log("No matches found");
      }

      await context.sync();
    } catch (error) {
      console.error("Error in replaceData:", error);
    }
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
    console.log("Action completed.");
  } catch (error) {
    console.error(error);
  }
}
