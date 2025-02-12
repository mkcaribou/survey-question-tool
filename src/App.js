import React from 'react';
import QuestionsDatabase from './QuestionsDatabase';

function App() {
  return (
    <div className="container mx-auto p-4">
      <h1 className="text-center text-2xl font-bold my-4">survey questions</h1>
      <QuestionsDatabase />
    </div>
  );
}

export default App;
