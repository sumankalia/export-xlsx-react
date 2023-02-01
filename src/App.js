import React, { useEffect, useState } from "react";
const ExcelJS = require("exceljs");

const toDataURL = (url) => {
  const promise = new Promise((resolve, reject) => {
    var xhr = new XMLHttpRequest();
    xhr.onload = function () {
      var reader = new FileReader();
      reader.readAsDataURL(xhr.response);
      reader.onloadend = function () {
        resolve({ base64Url: reader.result });
      };
    };
    xhr.open("GET", url);
    xhr.responseType = "blob";
    xhr.send();
  });

  return promise;
};

const App = () => {
  const [data, setData] = useState([]);
  useEffect(() => {
    fetch("https://dummyjson.com/products")
      .then((res) => res.json())
      .then(async (data) => {
        console.log(data);
        setData(data);
      })
      .then((json) => console.log(json));
  }, []);

  const exportExcelFile = () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("My Sheet");
    sheet.properties.defaultRowHeight = 80;

    sheet.getRow(1).border = {
      top: { style: "thick", color: { argb: "FFFF0000" } },
      left: { style: "thick", color: { argb: "000000FF" } },
      bottom: { style: "thick", color: { argb: "F08080" } },
      right: { style: "thick", color: { argb: "FF00FF00" } },
    };

    sheet.getRow(1).fill = {
      type: "pattern",
      pattern: "darkVertical",
      fgColor: { argb: "FFFF00" },
    };

    sheet.getRow(1).font = {
      name: "Comic Sans MS",
      family: 4,
      size: 16,
      bold: true,
    };

    sheet.columns = [
      {
        header: "Id",
        key: "id",
        width: 10,
      },
      { header: "Title", key: "title", width: 32 },
      {
        header: "Brand",
        key: "brand",
        width: 20,
      },
      {
        header: "Category",
        key: "category",
        width: 20,
      },
      {
        header: "Price",
        key: "price",
        width: 15,
      },
      {
        header: "Rating",
        key: "rating",
        width: 10,
      },
      {
        header: "Photo",
        key: "thumbnail",
        width: 30,
      },
    ];

    const promise = Promise.all(
      data?.products?.map(async (product, index) => {
        const rowNumber = index + 1;
        sheet.addRow({
          id: product?.id,
          title: product?.title,
          brand: product?.brand,
          category: product?.category,
          price: product?.price,
          rating: product?.rating,
        });
        console.log(product?.thumbnail);
        const result = await toDataURL(product?.thumbnail);
        const splitted = product?.thumbnail.split(".");
        const extName = splitted[splitted.length - 1];

        const imageId2 = workbook.addImage({
          base64: result.base64Url,
          extension: extName,
        });

        sheet.addImage(imageId2, {
          tl: { col: 6, row: rowNumber },
          ext: { width: 100, height: 100 },
        });
      })
    );

    promise.then(() => {
      const priceCol = sheet.getColumn(5);

      // iterate over all current cells in this column
      priceCol.eachCell((cell) => {
        const cellValue = sheet.getCell(cell?.address).value;
        // add a condition to set styling
        if (cellValue > 50 && cellValue < 1000) {
          sheet.getCell(cell?.address).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF0000" },
          };
        }
      });

      workbook.xlsx.writeBuffer().then(function (data) {
        const blob = new Blob([data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const url = window.URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = "download.xlsx";
        anchor.click();
        window.URL.revokeObjectURL(url);
      });
    });
  };

  return (
    <div style={{ padding: "30px" }}>
      <button
        className="btn btn-primary float-end mt-2 mb-2"
        onClick={exportExcelFile}
      >
        Export
      </button>
      <h3>Table Data:</h3>
      <table className="table table-bordered">
        <thead style={{ background: "yellow" }}>
          <tr>
            <th scope="col">Id</th>
            <th scope="col">Title</th>
            <th scope="col">Brand</th>
            <th scope="col">Category</th>
            <th scope="col">Price</th>
            <th scope="col">Rating</th>
            <th scope="col">Photo</th>
          </tr>
        </thead>
        <tbody>
          {Array.isArray(data?.products) &&
            data?.products?.map((row) => (
              <tr>
                <td>{row?.id}</td>
                <td>{row?.title}</td>
                <td>{row?.brand}</td>
                <td>{row?.category}</td>
                <td>${row?.price}</td>
                <td>{row?.rating}/5</td>
                <td>
                  <img src={row?.thumbnail} width="100" />
                </td>
              </tr>
            ))}
        </tbody>
      </table>
    </div>
  );
};

export default App;
