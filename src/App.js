import React, { useEffect, useState } from "react";
const ExcelJS = require('exceljs');

const App = () => {
  const [data, setData] = useState([]);
  useEffect(() => {
    fetch('https://dummyjson.com/products')
    .then(res => res.json())
    .then(data => {console.log(data); setData(data);
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('My Sheet');
      sheet.columns = [
        { header: 'Id', key: 'id', width: 10 },
        { header: 'Title', key: 'title', width: 32 },
        { header: 'Brand', key: 'brand', width: 20, outlineLevel: 1 },
        { header: 'Category', key: 'category', width: 20, outlineLevel: 1 },
        { header: 'Price', key: 'price', width: 15, outlineLevel: 1 },
        { header: 'Rating', key: 'rating', width: 10, outlineLevel: 1 },
        { header: 'Photo', key: 'thumbnail', width: 30, outlineLevel: 1 },
      ];

      data?.products?.map((product) => {
        sheet.addRow({id: product?.id, 
              title: product?.title,
              brand: product?.brand,
              category: product?.category,
              price: product?.price,
              rating: product?.rating,
            thumbnail: product?.thumbnail});
            });


      const priceCol = sheet.getColumn(5);

      // iterate over all current cells in this column
      priceCol.eachCell((cell) => {
      const cellValue = sheet.getCell(cell?.address).value;
        // add a condition to set styling
        if(cellValue > 50 && cellValue < 1000) {
          sheet.getCell(cell?.address).fill = {
            type: 'pattern',
            pattern:'solid',
            fgColor:{ argb:'FF0000' },
          };
        }
      });


      workbook.xlsx.writeBuffer().then(function (data) {
        const blob = new Blob([data],
          { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = 'download.xlsx';
        anchor.click();
        window.URL.revokeObjectURL(url);
      });
     })
    .then(json => console.log(json))
  }, []);

  return (
    <div style={{ padding: "30px" }}>

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
          </tr>
        </thead>
        <tbody>
          {Array.isArray(data?.products) && data?.products?.map((row) => (
            <tr>
              <td>{row?.id}</td>
              <td>{row?.title}</td>
              <td>{row?.brand}</td>
              <td>{row?.category}</td>
              <td>${row?.price}</td>
              <td>{row?.rating}/5</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default App;
