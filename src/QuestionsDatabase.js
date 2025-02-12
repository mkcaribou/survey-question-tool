import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { ChevronDown, ChevronUp, FileDown, FileText } from 'lucide-react';

const QuestionsDatabase = () => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  // filters available in the database
  const [filters, setFilters] = useState({
    categories: new Set(),
    subCategories: new Set(),
    questionTypes: new Set(),
    recommendedForStrive: new Set()
  });

  // selected filters for filtering data
  const [selectedFilters, setSelectedFilters] = useState({
    categories: new Set(),
    subCategories: new Set(),
    questionTypes: new Set(),
    recommendedForStrive: new Set()
  });

  // which filter sections are expanded
  const [expandedSections, setExpandedSections] = useState({
    categories: false,
    subCategories: false,
    questionTypes: false,
    recommendedForStrive: false
  });

  // selected rows by id
  const [selectedRowIds, setSelectedRowIds] = useState(new Set());

  // load data from the survey excel file located in public folder
  useEffect(() => {
    const loadData = async () => {
      try {
        setLoading(true);
        const surveyFileUrl = `${process.env.PUBLIC_URL}/static/survey.xlsx`;
	const res = await fetch(surveyFileUrl);
        if (!res.ok) {
          throw new Error('failed to fetch survey file');
        }
        const arrayBuffer = await res.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        if (!workbook.SheetNames.length) {
          throw new Error('no sheets found in workbook');
        }
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        let jsonData = XLSX.utils.sheet_to_json(firstSheet);
        if (!jsonData.length) {
          throw new Error('no data found in sheet');
        }
        // assign unique ids to each row
        jsonData = jsonData.map((row, index) => ({ id: index, ...row }));
        setData(jsonData);
        setFilters({
          categories: new Set(
            jsonData.map(row => row['Category'] || row.Category).filter(Boolean)
          ),
          subCategories: new Set(
            jsonData.map(row => row['Sub-category']).filter(Boolean)
          ),
          questionTypes: new Set(
            jsonData.map(row => row['Type of question']).filter(Boolean)
          ),
          recommendedForStrive: new Set(
            jsonData.map(row => row['Recommended for Strive']).filter(Boolean)
          )
        });
      } catch (err) {
        console.error('error loading data:', err);
        setError(err.message);
      } finally {
        setLoading(false);
      }
    };

    loadData();
  }, []);

  // memoized filtered data based on selected filters
  const filteredData = useMemo(() => {
    return data.filter(row => {
      const category = row['Category '] || row.Category;
      const subCategory = row['Sub-category'];
      const questionType = row['Type of question'];
      const recommended = row['Recommended for Strive'];

      const categoryMatch =
        selectedFilters.categories.size === 0 || selectedFilters.categories.has(category);
      const subCategoryMatch =
        selectedFilters.subCategories.size === 0 || selectedFilters.subCategories.has(subCategory);
      const questionTypeMatch =
        selectedFilters.questionTypes.size === 0 || selectedFilters.questionTypes.has(questionType);
      const recommendedMatch =
        selectedFilters.recommendedForStrive.size === 0 ||
        selectedFilters.recommendedForStrive.has(recommended);

      return categoryMatch && subCategoryMatch && questionTypeMatch && recommendedMatch;
    });
  }, [data, selectedFilters]);

  const toggleFilter = (filterKey, value) => {
    setSelectedFilters(prev => {
      const updated = new Set(prev[filterKey]);
      if (updated.has(value)) {
        updated.delete(value);
      } else {
        updated.add(value);
      }
      return { ...prev, [filterKey]: updated };
    });
  };

  const toggleSection = section => {
    setExpandedSections(prev => ({
      ...prev,
      [section]: !prev[section]
    }));
  };

  const toggleRowSelection = id => {
    setSelectedRowIds(prev => {
      const updated = new Set(prev);
      if (updated.has(id)) {
        updated.delete(id);
      } else {
        updated.add(id);
      }
      return updated;
    });
  };

  const handleSelectAll = e => {
    if (e.target.checked) {
      const allIds = filteredData.map(row => row.id);
      setSelectedRowIds(new Set(allIds));
    } else {
      // remove the filtered rows from the selection
      const newSelection = new Set(selectedRowIds);
      filteredData.forEach(row => newSelection.delete(row.id));
      setSelectedRowIds(newSelection);
    }
  };

  const exportToCSV = () => {
    const selectedData = filteredData.filter(row => selectedRowIds.has(row.id));
    if (!selectedData.length) return;

    const worksheet = XLSX.utils.json_to_sheet(selectedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Selected Questions');
    XLSX.writeFile(workbook, 'selected_questions.xlsx');
  };

  const exportToWord = () => {
    const selectedData = filteredData.filter(row => selectedRowIds.has(row.id));
    if (!selectedData.length) return;

    let docContent = '';
    selectedData.forEach((row, index) => {
      docContent += `question ${index + 1}: ${row.Question}\n`;
      docContent += `response options: ${row['Question response']}\n`;
      docContent += `type: ${row['Type of question']}\n`;
      docContent += `category: ${row['Category '] || row.Category}\n`;
      docContent += `sub-category: ${row['Sub-category']}\n`;
      docContent += `recommended for strive: ${row['Recommended for Strive']}\n\n`;
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
    return <div>loading survey questions...</div>;
  }
  if (error) {
    return <div>error loading data: {error}</div>;
  }

  const allFilteredSelected =
    filteredData.length > 0 && filteredData.every(row => selectedRowIds.has(row.id));

  return (
    <div className="p-4 space-y-4">
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        {/* filter section for categories */}
        <div className="border rounded-lg mb-4">
          <button
            onClick={() => toggleSection('categories')}
            className="w-full px-4 py-2 flex justify-between items-center bg-gray-50 rounded-t-lg hover:bg-gray-100"
          >
            <span>categories</span>
            {expandedSections.categories ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </button>
          {expandedSections.categories && (
            <div className="p-4 space-y-2">
              {Array.from(filters.categories).map(option => (
                <label key={option} className="flex items-center space-x-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={selectedFilters.categories.has(option)}
                    onChange={() => toggleFilter('categories', option)}
                    className="rounded border-gray-300"
                  />
                  <span>{option}</span>
                </label>
              ))}
            </div>
          )}
        </div>
        {/* filter section for sub-categories */}
        <div className="border rounded-lg mb-4">
          <button
            onClick={() => toggleSection('subCategories')}
            className="w-full px-4 py-2 flex justify-between items-center bg-gray-50 rounded-t-lg hover:bg-gray-100"
          >
            <span>sub-categories</span>
            {expandedSections.subCategories ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </button>
          {expandedSections.subCategories && (
            <div className="p-4 space-y-2">
              {Array.from(filters.subCategories).map(option => (
                <label key={option} className="flex items-center space-x-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={selectedFilters.subCategories.has(option)}
                    onChange={() => toggleFilter('subCategories', option)}
                    className="rounded border-gray-300"
                  />
                  <span>{option}</span>
                </label>
              ))}
            </div>
          )}
        </div>
        {/* filter section for question types */}
        <div className="border rounded-lg mb-4">
          <button
            onClick={() => toggleSection('questionTypes')}
            className="w-full px-4 py-2 flex justify-between items-center bg-gray-50 rounded-t-lg hover:bg-gray-100"
          >
            <span>question types</span>
            {expandedSections.questionTypes ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </button>
          {expandedSections.questionTypes && (
            <div className="p-4 space-y-2">
              {Array.from(filters.questionTypes).map(option => (
                <label key={option} className="flex items-center space-x-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={selectedFilters.questionTypes.has(option)}
                    onChange={() => toggleFilter('questionTypes', option)}
                    className="rounded border-gray-300"
                  />
                  <span>{option}</span>
                </label>
              ))}
            </div>
          )}
        </div>
        {/* filter section for recommended for strive */}
        <div className="border rounded-lg mb-4">
          <button
            onClick={() => toggleSection('recommendedForStrive')}
            className="w-full px-4 py-2 flex justify-between items-center bg-gray-50 rounded-t-lg hover:bg-gray-100"
          >
            <span>recommended for strive</span>
            {expandedSections.recommendedForStrive ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </button>
          {expandedSections.recommendedForStrive && (
            <div className="p-4 space-y-2">
              {Array.from(filters.recommendedForStrive).map(option => (
                <label key={option} className="flex items-center space-x-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={selectedFilters.recommendedForStrive.has(option)}
                    onChange={() => toggleFilter('recommendedForStrive', option)}
                    className="rounded border-gray-300"
                  />
                  <span>{option}</span>
                </label>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className="flex justify-between items-center mb-2">
        <div className="text-sm">
          showing {filteredData.length} of {data.length} questions
          {selectedRowIds.size > 0 && ` (${selectedRowIds.size} selected)`}
        </div>
      </div>

      <div className="overflow-x-auto border rounded-lg">
        <table className="min-w-full">
          <thead className="bg-gray-50">
            <tr>
              <th className="px-4 py-2 text-left w-8">
                <input
                  type="checkbox"
                  onChange={handleSelectAll}
                  checked={allFilteredSelected}
                  className="rounded border-gray-300"
                />
              </th>
              <th className="px-4 py-2 text-left">question</th>
              <th className="px-4 py-2 text-left w-64">response options</th>
              <th className="px-4 py-2 text-left">type</th>
              <th className="px-4 py-2 text-left">category</th>
              <th className="px-4 py-2 text-left">sub-category</th>
              <th className="px-4 py-2 text-left">recommended for strive</th>
            </tr>
          </thead>
          <tbody>
            {filteredData.map(row => (
              <tr key={row.id} className="hover:bg-gray-50">
                <td className="px-4 py-2">
                  <input
                    type="checkbox"
                    checked={selectedRowIds.has(row.id)}
                    onChange={() => toggleRowSelection(row.id)}
                    className="rounded border-gray-300"
                  />
                </td>
                <td className="px-4 py-2">{row.Question}</td>
                <td className="px-4 py-2">
                  <div className="max-h-24 overflow-y-auto">{row['Question response']}</div>
                </td>
                <td className="px-4 py-2">{row['Type of question']}</td>
                <td className="px-4 py-2">{row['Category '] || row.Category}</td>
                <td className="px-4 py-2">{row['Sub-category']}</td>
                <td className="px-4 py-2">{row['Recommended for Strive']}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="flex gap-4 justify-end">
        <button
          onClick={exportToCSV}
          className="px-4 py-2 bg-blue-500 text-white rounded disabled:opacity-50"
          disabled={selectedRowIds.size === 0}
        >
          <FileDown className="w-4 h-4 inline" /> export to excel
        </button>
        <button
          onClick={exportToWord}
          className="px-4 py-2 bg-blue-500 text-white rounded disabled:opacity-50"
          disabled={selectedRowIds.size === 0}
        >
          <FileText className="w-4 h-4 inline" /> export to word
        </button>
      </div>
    </div>
  );
};

export default QuestionsDatabase;
