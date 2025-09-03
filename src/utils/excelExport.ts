import * as XLSX from 'xlsx';
import { PrizeWinner } from '../types';

export const exportWinnersToExcel = (winners: PrizeWinner[]) => {
  // Prepare data for Excel export
  const excelData = winners.map((winner, index) => ({
    'S.No': index + 1,
    'Winner Name': winner.name,
    'Guide ID': winner.guide_id,
    'Department': winner.department,
    'Supervisor': winner.supervisor,
    'NPS Score': winner.nps,
    'NRPC Score': winner.nrpc,
    'Refund Percentage': winner.refund_percent,
    'Total Tickets': winner.total_tickets,
    'Prize Category': winner.prize_category || 'N/A',
    'Prize Name': winner.prize_name || 'N/A',
    'Won Date': new Date(winner.won_at).toLocaleDateString(),
    'Won Time': new Date(winner.won_at).toLocaleTimeString(),
    'Created Date': new Date(winner.created_at).toLocaleDateString(),
    'Winner ID': winner.id
  }));

  // Create workbook and worksheet
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(excelData);

  // Set column widths for better readability
  const columnWidths = [
    { wch: 8 },   // S.No
    { wch: 25 },  // Winner Name
    { wch: 12 },  // Guide ID
    { wch: 25 },  // Department
    { wch: 20 },  // Supervisor
    { wch: 12 },  // NPS Score
    { wch: 12 },  // NRPC Score
    { wch: 15 },  // Refund Percentage
    { wch: 12 },  // Total Tickets
    { wch: 25 },  // Prize Category
    { wch: 30 },  // Prize Name
    { wch: 12 },  // Won Date
    { wch: 12 },  // Won Time
    { wch: 12 },  // Created Date
    { wch: 40 }   // Winner ID
  ];
  worksheet['!cols'] = columnWidths;

  // Add worksheet to workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Winners Data');

  // Generate filename with current timestamp
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const filename = `Big_Dollar_Contest_Winners_${timestamp}.xlsx`;

  // Write and download the file
  XLSX.writeFile(workbook, filename);

  return {
    filename,
    recordCount: winners.length,
    exportDate: new Date().toLocaleString()
  };
};