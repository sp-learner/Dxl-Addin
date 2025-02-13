// src/components/Toolbar.js
import React from "react";

const Toolbar = () => {
  return (
    <>
      <div className="p-4 bg-gray-100">
        <h1 className="text-xl font-bold mb-4">DXL Add-In</h1>
        <div className="mb-6">
          <h2 className="text-lg font-semibold">Upload Data</h2>
          <button className="btn">Rapnet Upload</button>
          <button className="btn">Set Rapaport Rate</button>
          <button className="btn">Recut Planning</button>
        </div>

        <div>
          <h2 className="text-lg font-semibold">Tools</h2>
          <button className="btn">Custom Sort</button>
          <button className="btn">Demand Filter</button>
          <button className="btn">Custom Filter</button>
          <button className="btn">Format Sheet</button>
          <button className="btn">Put Formula</button>
        </div>
      </div>
    </>
  );
};

export default Toolbar;
