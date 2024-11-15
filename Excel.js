import { Injectable } from '@nestjs/common';
import * as ExcelJS from 'exceljs';
import { Response } from 'express';

@Injectable()
export class LoanReportService {
  async exportLoanStatusReport(response: Response): Promise<void> {
    // Tạo workbook và worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Báo cáo tình trạng giao dịch');

    // Thiết lập chiều rộng cho các cột
    worksheet.columns = [
      { width: 5 },   // Số TT
      { width: 25 },  // Loại hợp đồng vay
      { width: 20 },  // Số lượng tài khoản đã mở
      { width: 20 },  // Số lượng TK mở mới
      { width: 15 },  // Số lượng
      { width: 20 },  // Tổng giá trị
      { width: 20 },  // Giá trị giao dịch bình quân
      { width: 15 },  // Số lượng
      { width: 20 },  // Tổng giá trị
      { width: 20 },  // Giá trị giao dịch bình quân
      { width: 15 },  // Số lượng
      { width: 20 },  // Tổng giá trị
      { width: 20 },  // Giá trị giao dịch bình quân
      { width: 15 },  // Số lượng
      { width: 20 },  // Tổng giá trị
      { width: 20 },  // Giá trị giao dịch bình quân
    ];

    // Merge các ô cho phần header
    worksheet.mergeCells('A1:P1');
    worksheet.getCell('A1').value = 'BÁO CÁO TÌNH TRẠNG GIAO DỊCH';
    worksheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('A1').font = { size: 16, bold: true };

    worksheet.mergeCells('A2:P2');
    worksheet.getCell('A2').value = 'Từ: ____ đến ngày: ____';
    worksheet.getCell('A2').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('A2').font = { size: 12, italic: true };

    // Merge các ô cho phần header cột
    worksheet.mergeCells('A3:A4');
    worksheet.mergeCells('B3:B4');
    worksheet.mergeCells('C3:C4');
    worksheet.mergeCells('D3:D4');
    worksheet.mergeCells('E3:G3');
    worksheet.mergeCells('H3:J3');
    worksheet.mergeCells('K3:M3');
    worksheet.mergeCells('N3:P3');

    // Gán giá trị cho các ô header
    const headerCells = [
      'A3', 'B3', 'C3', 'D3', 'E3', 'H3', 'K3', 'N3',
      'E4', 'F4', 'G4', 'H4', 'I4', 'J4', 'K4', 'L4', 'M4', 'N4', 'O4', 'P4',
    ];

    const headers = [
      'S TT', 'Loại hợp đồng vay', 'Số lượng tài khoản đã mở đến ngày', 'Số lượng TK mở mới trong kỳ',
      'Giao dịch chờ thẩm định', 'Giao dịch chờ BCV giải ngân', 'Giao dịch đã tất toán', 'Giao dịch quá hạn',
      'Số lượng', 'Tổng giá trị', 'Giá trị giao dịch bình quân', 'Số lượng', 'Tổng giá trị',
      'Giá trị giao dịch bình quân', 'Số lượng', 'Tổng giá trị', 'Giá trị giao dịch bình quân',
      'Số lượng', 'Tổng giá trị', 'Giá trị giao dịch bình quân',
    ];

    // Đặt giá trị và border cho header
    headers.forEach((header, index) => {
      const cell = worksheet.getCell(headerCells[index]);
      cell.value = header;
      cell.font = { bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });

    // Thêm dữ liệu mẫu
    const data = [
      {
        type: 'Cho vay tín chấp',
        totalAccounts: 100,
        newAccounts: 10,
        pendingAmount: 5,
        pendingValue: 1000000,
        pendingAverage: 200000,
        disbursedAmount: 3,
        disbursedValue: 500000,
        disbursedAverage: 166667,
        settledAmount: 4,
        settledValue: 800000,
        settledAverage: 200000,
        overdueAmount: 2,
        overdueValue: 300000,
        overdueAverage: 150000,
      },
    ];

    // Thêm dữ liệu vào Excel
    let rowIndex = 5;
    data.forEach((item, index) => {
      const row = worksheet.addRow([
        index + 1,
        item.type,
        item.totalAccounts,
        item.newAccounts,
        item.pendingAmount,
        item.pendingValue,
        item.pendingAverage,
        item.disbursedAmount,
        item.disbursedValue,
        item.disbursedAverage,
        item.settledAmount,
        item.settledValue,
        item.settledAverage,
        item.overdueAmount,
        item.overdueValue,
        item.overdueAverage,
      ]);

      // Thiết lập border và wrapText cho các ô có dữ liệu
      row.eachCell((cell) => {
        cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });

      rowIndex++;
    });

    // Thiết lập header để tải file
    response.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    response.setHeader(
      'Content-Disposition',
      'attachment; filename=LoanReport.xlsx',
    );

    // Xuất file Excel
    await workbook.xlsx.write(response);
    response.end();
  }
}
