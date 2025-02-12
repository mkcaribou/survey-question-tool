import React from 'react';
import QuestionsDatabase from './QuestionsDatabase';

function App() {
  return (
    <div className="min-h-screen flex flex-col items-center justify-start bg-gray-100">
      <h1 className="text-4xl font-bold text-blue-600 mt-4">Survey Questions Database</h1>
      <QuestionsDatabase />
    </div>
  );
}

export default App;
