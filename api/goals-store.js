import fs from 'fs';
import path from 'path';

// En memoria para Vercel serverless (fallback temporal)
let inMemoryGoals = {};

const getGoalsFilePath = () => {
  // En local, guardar en el directorio de la API. En Vercel, en /tmp.
  const isVercel = process.env.VERCEL || process.env.NOW_BUILDER;
  if (isVercel) {
    return '/tmp/goals-db.json';
  }
  return path.join(process.cwd(), 'api', 'goals-db.json');
};

const loadLocalGoals = () => {
  try {
    const filePath = getGoalsFilePath();
    if (fs.existsSync(filePath)) {
      const data = fs.readFileSync(filePath, 'utf8');
      return JSON.parse(data);
    }
  } catch (e) {
    console.error('Error loading local goals', e);
  }
  return inMemoryGoals;
};

const saveLocalGoals = (goals) => {
  try {
    const filePath = getGoalsFilePath();
    fs.writeFileSync(filePath, JSON.stringify(goals, null, 2), 'utf8');
  } catch (e) {
    console.error('Error saving local goals', e);
  }
  inMemoryGoals = goals;
};

// Utilidad para Vercel KV
const useKV = () => {
  return !!process.env.KV_REST_API_URL;
};

const getKVHeaders = () => {
  return {
    'Authorization': `Bearer ${process.env.KV_REST_API_TOKEN}`,
    'Content-Type': 'application/json'
  };
};

export async function getGoal(period) {
  if (useKV()) {
    try {
      const url = `${process.env.KV_REST_API_URL}/get/finance_goal_${period}`;
      const response = await fetch(url, { headers: getKVHeaders() });
      if (response.ok) {
        const payload = await response.json();
        const value = parseFloat(payload.result);
        return Number.isFinite(value) ? value : 0;
      }
    } catch (e) {
      console.error('Error getting goal from KV', e);
    }
  }
  
  const goals = loadLocalGoals();
  return parseFloat(goals[period]) || 0;
}

export async function setGoal(period, value) {
  const numValue = Math.max(0, parseFloat(value) || 0);
  
  if (useKV()) {
    try {
      const url = `${process.env.KV_REST_API_URL}/set/finance_goal_${period}/${numValue}`;
      const response = await fetch(url, { headers: getKVHeaders() });
      if (response.ok) {
        return true;
      }
    } catch (e) {
      console.error('Error setting goal in KV', e);
    }
  }
  
  const goals = loadLocalGoals();
  goals[period] = numValue;
  saveLocalGoals(goals);
  return true;
}
