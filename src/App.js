import React from 'react';
import QuestionsDatabase from './QuestionsDatabase';

function App() {
  return (
    <div className="min-h-screen flex flex-col items-center justify-center bg-gray-100">
      <h1 className="text-4xl font-bold text-blue-600 mb-4">Survey Questions Database</h1>
      {/* Test element */}
      <div className="w-32 h-32 bg-red-500 text-white flex items-center justify-center">
        Test
      </div>
      <QuestionsDatabase />
    </div>
  );
}

export default App;
