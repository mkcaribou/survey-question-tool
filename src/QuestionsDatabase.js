import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { ChevronDown, ChevronUp, FileDown, FileText } from 'lucide-react';

const QuestionsDatabase = () => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  // Available filters based on the loaded data
  const [filters, setFilters] = useState({
    categories: new Set(),
    subCategories: new Set(),
    questionTypes: new Set(),
    recommendedForStrive: new Set()
  });

  // User-selected filters
  const [selectedFilters, setSelectedFilters] = useState({
    categories: new Set(),
    subCategories: new Set(),
    questionTypes: new Set(),
    recommendedForStrive: new Set()
  });

  // Which filter dropdown sections are expanded
  const [expandedSections, setExpandedSections] = useState({
    categories: false,
    subCategories: false,
    questionTypes: false,
    recommendedForStrive: false
  });

  // Track selected row IDs for export
  const [selectedRowIds, setSelectedRowIds] = useState(new Set());

  useEffect(() => {
    const loadData = async () => {
      try {
        setLoading(true);
        const surveyFileUrl = `${process.env.PUBLIC_URL}/survey.xlsx`;
        const res = await fetch(surveyFileUrl);
        if (!res.ok) {
          throw new Error(`Failed to fetch survey file. Status: ${res.status}`);
        }
        const arrayBuffer = await res.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        if (!workbook.SheetNames.length) {
          throw new Error('No sheets found in workbook');
        }
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        let jsonData = XLSX.utils.sheet_to_json(firstSheet);
        if (!jsonData.length) {
          throw new Error('No data found in sheet');
        }
        // Trim any extra whitespace from keys
        jsonData = jsonData.map((row, index) => {
          const trimmedRow = {};
          Object.keys(row).forEach((key) => {
            trimmedRow[key.trim()] = row[key];
          });
          return { id: index, ...trimmedRow };
        });
        setData(jsonData);
        setFilters({
          categories: new Set(jsonData.map((row) => row.Category).filter(Boolean)),
          subCategories: new Set(jsonData.map((row) => row['Sub-category']).filter(Boolean)),
          questionTypes: new Set(jsonData.map((row) => row['Type of question']).filter(Boolean)),
          recommendedForStrive: new Set(jsonData.map((row) => row['Recommended for Strive']).filter(Boolean))
        });
      } catch (err) {
        console.error('Error loading data:', err);
        setError(err.message);
      } finally {
        setLoading(false);
      }
    };

    loadData();
  }, []);

  const filteredData = useMemo(() => {
    return data.filter((row) => {
      const categoryMatch =
        selectedFilters.categories.size === 0 || selectedFilters.categories.has(row.Category);
      const subCategoryMatch =
        selectedFilters.subCategories.size === 0 || selectedFilters.subCategories.has(row['Sub-category']);
      const questionTypeMatch =
        selectedFilters.questionTypes.size === 0 || selectedFilters.questionTypes.has(row['Type of question']);
      const recommendedMatch =
        selectedFilters.recommendedForStrive.size === 0 ||
        selectedFilters.recommendedForStrive.has(row['Recommended for Strive']);
      return categoryMatch && subCategoryMatch && questionTypeMatch && recommendedMatch;
    });
  }, [data, selectedFilters]);

  const toggleFilter = (filterKey, value) => {
    setSelectedFilters((prev) => {
      const updated = new Set(prev[filterKey]);
      if (updated.has(value)) {
        updated.delete(value);
      } else {
        updated.add(value);
      }
      return { ...prev, [filterKey]: updated };
    });
  };

  const toggleSection = (section) => {
    setExpandedSections((prev) => ({
      ...prev,
      [section]: !prev[section]
    }));
  };

  const toggleRowSelection = (id) => {
    setSelectedRowIds((prev) => {
      const updated = new Set(prev);
      if (updated.has(id)) {
        updated.delete(id);
      } else {
        updated.add(id);
      }
      return updated;
    });
  };

  const handleSelectAll = (e) => {
    if (e.target.checked) {
      const allIds = filteredData.map((row) => row.id);
      setSelectedRowIds(new Set(allIds));
    } else {
      // Remove the filtered rows from the selection
      const newSelection = new Set(selectedRowIds);
      filteredData.forEach((row) => newSelection.delete(row.id));
      setSelectedRowIds(newSelection);
    }
  };

  const clearFilters = () => {
    setSelectedFilters({
      categories: new Set(),
      subCategories: new Set(),
      questionTypes: new Set(),
      recommendedForStrive: new Set()
    });
  };

  const exportToCSV = () => {
    const selectedData = filteredData.filter((row) => selectedRowIds.has(row.id));
    if (!selectedData.length) return;

    const worksheet = XLSX.utils.json_to_sheet(selectedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Selected Questions');
    XLSX.writeFile(workbook, 'selected_questions.xlsx');
  };

  const exportToWord = () => {
    const selectedData = filteredData.filter((row) => selectedRowIds.has(row.id));
    if (!selectedData.length) return;

    let docContent = '';
    selectedData.forEach((row, index) => {
      docContent += `Question ${index + 1}: ${row.Question}\n`;
      docContent += `Response Options: ${row['Question response']}\n`;
      docContent += `Type: ${row['Type of question']}\n`;
      docContent += `Category: ${row.Category}\n`;
      docContent += `Sub-category: ${row['Sub-category']}\n`;
      docContent += `Recommended for Strive: ${row['Recommended for Strive']}\n\n`;
    });

    const blob = new Blob([docContent], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'selected_questions.doc';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  if (loading) {
    return <div className="p-6 text-center text-gray-600">Loading survey questions...</div>;
  }
  if (error) {
    return <div className="p-6 text-center text-red-600">Error loading data: {error}</div>;
  }

  const allFilteredSelected =
    filteredData.length > 0 && filteredData.every((row) => selectedRowIds.has(row.id));

  return (
    <div className="p-6 max-w-7xl mx-auto">
      <h1 className="text-3xl font-bold text-center mb-6">Survey Questions</h1>
      
      {/* Filters Section */}
      <div className="mb-6">
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          {/* Categories Filter */}
          <div className="border rounded shadow">
            <button
              onClick={() => toggleSection('categories')}
              className="w-full px-4 py-2 bg-blue-100 hover:bg-blue-200 text-left font-semibold flex justify-between items-center focus:outline-none"
              aria-expanded={expandedSections.categories}
            >
              Categories
              {expandedSections.categories ? (
                <ChevronUp className="w-5 h-5" />
              ) : (
                <ChevronDown className="w-5 h-5" />
              )}
            </button>
            {expandedSections.categories && (
              <div className="p-4">
                {Array.from(filters.categories).map((option) => (
                  <label key={option} className="flex items-center space-x-2">
                    <input
                      type="checkbox"
                      checked={selectedFilters.categories.has(option)}
                      onChange={() => toggleFilter('categories', option)}
                      className="rounded border-gray-300 focus:ring-blue-500"
                    />
                    <span>{option}</span>
                  </label>
                ))}
              </div>
            )}
          </div>
          
          {/* Sub-Categories Filter */}
          <div className="border rounded shadow">
            <button
              onClick={() => toggleSection('subCategories')}
              className="w-full px-4 py-2 bg-blue-100 hover:bg-blue-200 text-left font-semibold flex justify-between items-center focus:outline-none"
              aria-expanded={expandedSections.subCategories}
            >
              Sub-Categories
              {expandedSections.subCategories ? (
                <ChevronUp className="w-5 h-5" />
              ) : (
                <ChevronDown className="w-5 h-5" />
              )}
            </button>
            {expandedSections.subCategories && (
              <div className="p-4">
                {Array.from(filters.subCategories).map((option) => (
                  <label key={option} className="flex items-center space-x-2">
                    <input
                      type="checkbox"
                      checked={selectedFilters.subCategories.has(option)}
                      onChange={() => toggleFilter('subCategories', option)}
                      className="rounded border-gray-300 focus:ring-blue-500"
                    />
                    <span>{option}</span>
                  </label>
                ))}
              </div>
            )}
          </div>
          
          {/* Question Types Filter */}
          <div className="border rounded shadow">
            <button
              onClick={() => toggleSection('questionTypes')}
              className="w-full px-4 py-2 bg-blue-100 hover:bg-blue-200 text-left font-semibold flex justify-between items-center focus:outline-none"
              aria-expanded={expandedSections.questionTypes}
            >
              Question Types
              {expandedSections.questionTypes ? (
                <ChevronUp className="w-5 h-5" />
              ) : (
                <ChevronDown className="w-5 h-5" />
              )}
            </button>
            {expandedSections.questionTypes && (
              <div className="p-4">
                {Array.from(filters.questionTypes).map((option) => (
                  <label key={option} className="flex items-center space-x-2">
                    <input
                      type="checkbox"
                      checked={selectedFilters.questionTypes.has(option)}
                      onChange={() => toggleFilter('questionTypes', option)}
                      className="rounded border-gray-300 focus:ring-blue-500"
                    />
                    <span>{option}</span>
                  </label>
                ))}
              </div>
            )}
          </div>
          
          {/* Recommended for Strive Filter */}
          <div className="border rounded shadow">
            <button
              onClick={() => toggleSection('recommendedForStrive')}
              className="w-full px-4 py-2 bg-blue-100 hover:bg-blue-200 text-left font-semibold flex justify-between items-center focus:outline-none"
              aria-expanded={expandedSections.recommendedForStrive}
            >
              Recommended for Strive
              {expandedSections.recommendedForStrive ? (
                <ChevronUp className="w-5 h-5" />
              ) : (
                <ChevronDown className="w-5 h-5" />
              )}
            </button>
            {expandedSections.recommendedForStrive && (
              <div className="p-4">
                {Array.from(filters.recommendedForStrive).map((option) => (
                  <label key={option} className="flex items-center space-x-2">
                    <input
                      type="checkbox"
                      checked={selectedFilters.recommendedForStrive.has(option)}
                      onChange={() => toggleFilter('recommendedForStrive', option)}
                      className="rounded border-gray-300 focus:ring-blue-500"
                    />
                    <span>{option}</span>
                  </label>
                ))}
              </div>
            )}
          </div>
        </div> 
        
        {/* Clear Options Button */}
        <div className="mt-4 text-right">
          <button
            onClick={clearFilters}
            className="px-4 py-2 bg-red-500 text-white rounded hover:bg-red-600 focus:outline-none focus:ring-2 focus:ring-red-300"
          >
            Clear Options
          </button>
        </div>
      </div>
      
      {/* Table Section */}
      <div className="overflow-x-auto">
        <table className="min-w-full bg-white border border-gray-200 rounded shadow">
          <thead className="bg-blue-500 text-white">
            <tr>
              <th className="px-4 py-2">
                <input
                  type="checkbox"
                  onChange={handleSelectAll}
                  checked={allFilteredSelected}
                  className="rounded border-gray-300 focus:ring-blue-500"
                  aria-label="Select all questions"
                />
              </th>
              <th className="px-4 py-2">Question</th>
              <th className="px-4 py-2">Response Options</th>
              <th className="px-4 py-2">Type</th>
              <th className="px-4 py-2">Category</th>
              <th className="px-4 py-2">Sub-Category</th>
              <th className="px-4 py-2">Recommended for Strive</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-200">
            {filteredData.map((row) => (
              <tr key={row.id} className="hover:bg-gray-100">
                <td className="px-4 py-2 text-center">
                  <input
                    type="checkbox"
                    checked={selectedRowIds.has(row.id)}
                    onChange={() => toggleRowSelection(row.id)}
                    className="rounded border-gray-300 focus:ring-blue-500"
                    aria-label={`Select question: ${row.Question}`}
                  />
                </td>
                <td className="px-4 py-2">{row.Question}</td>
                <td className="px-4 py-2">
                  <div className="max-h-24 overflow-y-auto">{row['Question response']}</div>
                </td>
                <td className="px-4 py-2">{row['Type of question']}</td>
                <td className="px-4 py-2">{row.Category}</td>
                <td className="px-4 py-2">{row['Sub-category']}</td>
                <td className="px-4 py-2">{row['Recommended for Strive']}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      
      {/* Export Buttons */}
      <div className="flex gap-4 justify-end mt-6">
        <button
          onClick={exportToCSV}
          className="flex items-center gap-2 px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 focus:outline-none focus:ring-2 focus:ring-green-300"
          disabled={selectedRowIds.size === 0}
        >
          <FileDown className="w-5 h-5" />
          Export to Excel
        </button>
        <button
          onClick={exportToWord}
          className="flex items-center gap-2 px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 focus:outline-none focus:ring-2 focus:ring-green-300"
          disabled={selectedRowIds.size === 0}
        >
          <FileText className="w-5 h-5" />
          Export to Word
        </button>
      </div>
    </div>
  );
};

export default QuestionsDatabase;
