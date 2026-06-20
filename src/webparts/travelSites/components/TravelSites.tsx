import * as React from 'react';
import { useState, useEffect } from 'react';
import { 
  DetailsList, 
  SelectionMode, 
  IColumn, 
  TextField, 
  Dropdown, 
  IDropdownOption, 
  PrimaryButton, 
  DefaultButton, 
  Stack, 
  MessageBar, 
  MessageBarType 
} from '@fluentui/react';

import * as XLSX from 'xlsx';

// Safe ESM compilation imports for pdfmake
import pdfMake from 'pdfmake/build/pdfmake';
import pdfFonts from 'pdfmake/build/vfs_fonts';

export interface ITravelSiteItem {
  Id?: number;
  Title: string; 
  Place: string;
  FamousFood: string;
  Price: 'Budget-friendly' | 'Premium Experience' | 'No cost' | 'Mid-range';
}

export interface ITravelSitesProps {
  context: any;
}

const ITEMS_PER_PAGE = 2; 

export const TravelSites = (props: ITravelSitesProps): JSX.Element => {
  const [items, setItems] = useState<ITravelSiteItem[]>([
    { Id: 1, Title: 'Japan', Place: 'Kyoto', FamousFood: 'Ramen', Price: 'Mid-range' },
    { Id: 2, Title: 'Italy', Place: 'Rome', FamousFood: 'Pizza Margherita', Price: 'Budget-friendly' },
    { Id: 3, Title: 'France', Place: 'Paris', FamousFood: 'Croissants', Price: 'Premium Experience' },
    { Id: 4, Title: 'India', Place: 'Taj Mahal', FamousFood: 'Biryani', Price: 'Budget-friendly' },
    { Id: 5, Title: 'Egypt', Place: 'Cairo', FamousFood: 'Koshary', Price: 'No cost' },
    { Id: 6, Title: 'Spain', Place: 'Barcelona', FamousFood: 'Paella', Price: 'Mid-range' }
  ]);

  const [filteredItems, setFilteredItems] = useState<ITravelSiteItem[]>([]);
  const [paginatedItems, setPaginatedItems] = useState<ITravelSiteItem[]>([]);
  const [filterText, setFilterText] = useState<string>('');

  const [currentPage, setCurrentPage] = useState<string>('All'); 
  const [pageOptions, setPageOptions] = useState<IDropdownOption[]>([]);

  const [sortColumnKey, setSortColumnKey] = useState<string>('');
  const [isSortedDescending, setIsSortedDescending] = useState<boolean>(false);
  
  const [itemId, setItemId] = useState<number | null>(null);
  const [destination, setDestination] = useState<string>('');
  const [place, setPlace] = useState<string>('');
  const [famousFood, setFamousFood] = useState<string>('');
  const [price, setPrice] = useState<any>('Mid-range');
  const [message, setMessage] = useState<{ text: string; type: MessageBarType } | null>(null);

  const priceOptions: IDropdownOption[] = [
    { key: 'Budget-friendly', text: 'Budget-friendly' },
    { key: 'Premium Experience', text: 'Premium Experience' },
    { key: 'No cost', text: 'No cost' },
    { key: 'Mid-range', text: 'Mid-range' },
  ];

  useEffect(() => {
    let result = [...items];

    if (filterText !== '') {
      const lowercasedFilter = filterText.toLowerCase();
      result = result.filter(item => 
        (item.Title && item.Title.toLowerCase().includes(lowercasedFilter)) ||
        (item.Place && item.Place.toLowerCase().includes(lowercasedFilter)) ||
        (item.FamousFood && item.FamousFood.toLowerCase().includes(lowercasedFilter)) ||
        (item.Price && item.Price.toLowerCase().includes(lowercasedFilter))
      );
    }

    if (sortColumnKey) {
      result.sort((a, b) => {
        const valA = (a[sortColumnKey as keyof ITravelSiteItem] || '').toString().toLowerCase();
        const valB = (b[sortColumnKey as keyof ITravelSiteItem] || '').toString().toLowerCase();
        return isSortedDescending ? valB.localeCompare(valA) : valA.localeCompare(valB);
      });
    }

    setFilteredItems(result);
    setCurrentPage('All'); 
  }, [filterText, items, sortColumnKey, isSortedDescending]);

  useEffect(() => {
    const totalPages = Math.ceil(filteredItems.length / ITEMS_PER_PAGE);
    const options: IDropdownOption[] = [{ key: 'All', text: 'All Pages' }];
    
    for (let i = 1; i <= totalPages; i++) {
      options.push({ key: i.toString(), text: `Page ${i}` });
    }
    setPageOptions(options);

    if (currentPage === 'All') {
      setPaginatedItems(filteredItems);
    } else {
      const pageNum = parseInt(currentPage, 10);
      const startIndex = (pageNum - 1) * ITEMS_PER_PAGE;
      const endIndex = startIndex + ITEMS_PER_PAGE;
      setPaginatedItems(filteredItems.slice(startIndex, endIndex));
    }
  }, [filteredItems, currentPage]);

  const showMsg = (text: string, type: MessageBarType) => {
    setMessage({ text, type });
    setTimeout(() => setMessage(null), 4000);
  };

  const clearForm = () => {
    setItemId(null);
    setDestination('');
    setPlace('');
    setFamousFood('');
    setPrice('Mid-range');
  };

  const handleSave = () => {
    if (!destination || !place) {
      showMsg("Destination and Place fields are required.", MessageBarType.warning);
      return;
    }

    if (itemId) {
      const updatedItems = items.map(item => item.Id === itemId ? {
        Id: itemId, Title: destination, Place: place, FamousFood: famousFood, Price: price
      } : item);
      setItems(updatedItems);
      showMsg("Record updated successfully!", MessageBarType.success);
    } else {
      const newItem: ITravelSiteItem = {
        Id: Math.floor(Math.random() * 100000),
        Title: destination, Place: place, FamousFood: famousFood, Price: price
      };
      setItems([...items, newItem]);
      showMsg("Record added successfully!", MessageBarType.success);
    }
    clearForm();
  };

  const handleEdit = (item: ITravelSiteItem) => {
    setItemId(item.Id || null);
    setDestination(item.Title);
    setPlace(item.Place);
    setFamousFood(item.FamousFood);
    setPrice(item.Price);
  };

  const handleDelete = (id: number) => {
    if (window.confirm("Are you sure you want to delete this record?")) {
      setItems(items.filter(item => item.Id !== id));
      showMsg("Record removed.", MessageBarType.error);
    }
  };

  const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const isDescending = column.key === sortColumnKey ? !isSortedDescending : false;
    setSortColumnKey(column.key);
    setIsSortedDescending(isDescending);
  };

  const getTargetExportData = (): ITravelSiteItem[] => {
    return currentPage === 'All' ? filteredItems : paginatedItems;
  };

  const exportToExcel = (): void => {
    const targetData = getTargetExportData();
    const exportRows = targetData.map(item => ({
      'Destination': item.Title,
      'Place': item.Place,
      'Famous Food': item.FamousFood,
      'Price Profile': item.Price
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Travel Data");
    
    XLSX.writeFile(workbook, `TravelSites_Export_${currentPage}.xlsx`);
    showMsg(`Exported scope [${currentPage}] to Excel successfully!`, MessageBarType.success);
  };

  const exportToPDF = (): void => {
    const targetData = getTargetExportData();
    const scopeLabel = currentPage === 'All' ? 'All Pages' : `Page ${currentPage}`;

    const tableBody = [
      [
        { text: 'Destination', bold: true, fillColor: '#e81123', color: '#ffffff', margin: [0, 4, 0, 4] },
        { text: 'Place', bold: true, fillColor: '#e81123', color: '#ffffff', margin: [0, 4, 0, 4] },
        { text: 'Famous Food', bold: true, fillColor: '#e81123', color: '#ffffff', margin: [0, 4, 0, 4] },
        { text: 'Price Profile', bold: true, fillColor: '#e81123', color: '#ffffff', margin: [0, 4, 0, 4] }
      ]
    ];

    targetData.forEach(item => {
      tableBody.push([
        { text: item.Title || '', bold: false, fillColor: '', color: '#333333', margin: [0, 3, 0, 3] },
        { text: item.Place || '', bold: false, fillColor: '', color: '#333333', margin: [0, 3, 0, 3] },
        { text: item.FamousFood || '', bold: false, fillColor: '', color: '#333333', margin: [0, 3, 0, 3] },
        { text: item.Price || '', bold: false, fillColor: '', color: '#333333', margin: [0, 3, 0, 3] }
      ]);
    });

    const documentDefinition: any = {
      content: [
        { text: 'Travel Sites Directory', style: 'header' },
        { text: `Export Configuration Scope: ${scopeLabel}`, style: 'subheader' },
        { text: `Total Exported Rows: ${targetData.length}`, margin: [0, 0, 0, 15], fontSize: 10, italic: true },
        {
          table: {
            headerRows: 1,
            widths: ['25%', '25%', '25%', '25%'],
            body: tableBody
          },
          layout: {
            hLineWidth: (i: number, node: any) => (i === 0 || i === node.table.body.length) ? 1 : 0.5,
            vLineWidth: () => 0.5,
            hLineColor: () => '#dddddd',
            vLineColor: () => '#dddddd'
          }
        }
      ],
      styles: {
        header: { fontSize: 18, bold: true, margin: [0, 0, 0, 5], color: '#222222' },
        subheader: { fontSize: 12, bold: false, margin: [0, 0, 0, 2], color: '#666666' }
      }
    };

    // Safe global key extraction strategy to fully bypass ts(2339) 
    const fontTarget = pdfFonts && (pdfFonts as any).pdfMake ? (pdfFonts as any).pdfMake.vfs : (pdfMake as any).vfs;
    (pdfMake as any).vfs = fontTarget;
    
    pdfMake.createPdf(documentDefinition).download(`TravelSites_Directory_${currentPage}.pdf`);
    showMsg(`Exported scope [${currentPage}] to PDF successfully!`, MessageBarType.success);
  };

  const columns: IColumn[] = [
    { 
      key: 'Title', name: 'Destination', fieldName: 'Title', minWidth: 100, maxWidth: 140, isResizable: true,
      isSorted: sortColumnKey === 'Title', isSortedDescending: isSortedDescending, onColumnClick: onColumnClick 
    },
    { 
      key: 'Place', name: 'Place', fieldName: 'Place', minWidth: 100, maxWidth: 140, isResizable: true,
      isSorted: sortColumnKey === 'Place', isSortedDescending: isSortedDescending, onColumnClick: onColumnClick 
    },
    { 
      key: 'FamousFood', name: 'Famous Food', fieldName: 'FamousFood', minWidth: 110, maxWidth: 140, isResizable: true,
      isSorted: sortColumnKey === 'FamousFood', isSortedDescending: isSortedDescending, onColumnClick: onColumnClick 
    },
    { 
      key: 'Price', name: 'Price Profile', fieldName: 'Price', minWidth: 110, maxWidth: 130, isResizable: true,
      isSorted: sortColumnKey === 'Price', isSortedDescending: isSortedDescending, onColumnClick: onColumnClick 
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 160,
      onRender: (item: ITravelSiteItem): JSX.Element => (
        <Stack horizontal tokens={{ childrenGap: 6 }}>
          <DefaultButton text="Edit" onClick={() => handleEdit(item)} iconProps={{ iconName: 'Edit' }} />
          <DefaultButton text="Delete" onClick={() => handleDelete(item.Id!)} iconProps={{ iconName: 'Delete' }} styles={{ root: { color: '#a80000' } }} />
        </Stack>
      )
    }
  ];

  return (
    <Stack tokens={{ childrenGap: 20 }} style={{ padding: 20, backgroundColor: '#fff', borderRadius: 4 }}>
      <h2>Travel Sites Directory Matrix</h2>
      
      {message && <MessageBar messageBarType={message.type}>{message.text}</MessageBar>}

      <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end">
        <TextField label="Destination" value={destination} onChange={(e, val) => setDestination(val || '')} required />
        <TextField label="Place" value={place} onChange={(e, val) => setPlace(val || '')} required />
        <TextField label="Famous Food" value={famousFood} onChange={(e, val) => setFamousFood(val || '')} />
        <Dropdown 
          label="Price Profile" 
          options={priceOptions} 
          selectedKey={price} 
          onChange={(e, option) => setPrice(option?.key as any)} 
          style={{ width: 160 }}
        />
        <Stack horizontal tokens={{ childrenGap: 6 }}>
          <PrimaryButton text={itemId ? "Update" : "Add"} onClick={handleSave} />
          {itemId && <DefaultButton text="Cancel" onClick={clearForm} />}
        </Stack>
      </Stack>

      <hr style={{ border: '0.5px solid #eee', margin: '5px 0' }} />

      <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="end" style={{ backgroundColor: '#fafafa', padding: 10, borderRadius: 2 }}>
        <Dropdown 
          label="-- select page --" 
          options={pageOptions} 
          selectedKey={currentPage} 
          onChange={(e, opt) => setCurrentPage(opt?.key as string)} 
          style={{ width: 180 }}
        />
        
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <DefaultButton 
            text={`Export ${currentPage === 'All' ? 'All Data' : 'This Page'} to Excel`} 
            iconProps={{ iconName: 'ExcelDocument' }} 
            onClick={exportToExcel}
            styles={{
              root: { backgroundColor: '#107c41', color: '#fff', border: 'none', borderRadius: 2 },
              rootHovered: { backgroundColor: '#0b592e', color: '#fff' },
              icon: { color: '#fff' }
            }}
          />
          <DefaultButton 
            text={`Export ${currentPage === 'All' ? 'All Data' : 'This Page'} to PDF`} 
            iconProps={{ iconName: 'PDF' }} 
            onClick={exportToPDF}
            styles={{
              root: { backgroundColor: '#e81123', color: '#fff', border: 'none', borderRadius: 2 },
              rootHovered: { backgroundColor: '#b30b18', color: '#fff' },
              icon: { color: '#fff' }
            }}
          />
        </Stack>
      </Stack>

      <Stack horizontal style={{ maxWidth: 400 }}>
        <TextField 
          label="Filter Table Search" 
          placeholder="Type parameter keywords..." 
          value={filterText} 
          onChange={(e, val) => setFilterText(val || '')} 
          iconProps={{ iconName: 'Filter' }}
          style={{ width: '100%' }}
        />
      </Stack>

      <DetailsList
        items={paginatedItems}
        columns={columns}
        selectionMode={SelectionMode.none}
      />
    </Stack>
  );
};