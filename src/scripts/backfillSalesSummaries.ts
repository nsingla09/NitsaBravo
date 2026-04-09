import { db } from '../firebase';
import { collection, getDocs, writeBatch, doc, serverTimestamp } from 'firebase/firestore';

/**
 * Backfill Sales Summaries
 * 
 * This script reads all existing sales and populates the salesSummaries collection.
 * Run this once after deploying the new summary-based dashboard.
 */
export async function backfillSalesSummaries() {
  console.log("Starting backfill...");
  const salesSnap = await getDocs(collection(db, 'sales'));
  const sales = salesSnap.docs.map(d => ({ id: d.id, ...d.data() }));
  
  const summaries: any = {};

  function addtoSummary(scope, type, periodId, revenue, status) {
    const key = `${type}_${scope}_${periodId}`;
    if (!summaries[key]) {
      const scopeParts = scope.split('_');
      summaries[key] = {
        totalRevenue: 0,
        totalSalesCount: 0,
        statusCounts: {},
        statusValues: {},
        type,
        scope: scope === 'global' ? 'global' : scopeParts[0],
        scopeId: scope === 'global' ? 'global' : scopeParts[1],
        periodId,
        lastUpdatedAt: serverTimestamp()
      };
    }
    
    summaries[key].totalRevenue += revenue;
    summaries[key].totalSalesCount += 1;
    if (status) {
      summaries[key].statusCounts[status] = (summaries[key].statusCounts[status] || 0) + 1;
      summaries[key].statusValues[status] = (summaries[key].statusValues[status] || 0) + revenue;
    }
  }

  sales.forEach((sale: any) => {
    const revenue = sale.packageValue || 0;
    const status = sale.advanceCN;
    const date = sale.date;
    const week = sale.week;
    const bde = sale.bde;
    const agent = sale.agent;
    
    // Month extraction
    const d = new Date(date);
    const month = d.toLocaleString('default', { month: 'long', year: 'numeric' });

    const scopes = ['global'];
    if (bde) scopes.push(`bde_${bde}`);
    if (agent) scopes.push(`agent_${agent}`);

    scopes.forEach(scope => {
      addtoSummary(scope, 'daily', date, revenue, status);
      addtoSummary(scope, 'weekly', week, revenue, status);
      addtoSummary(scope, 'monthly', month, revenue, status);
      addtoSummary(scope, 'lifetime', 'total', revenue, status);
    });
  });

  // Write to Firestore in batches
  const batch = writeBatch(db);
  Object.keys(summaries).forEach(key => {
    const ref = doc(db, 'salesSummaries', key);
    batch.set(ref, summaries[key]);
  });

  await batch.commit();
  console.log(`Backfill complete. Updated ${Object.keys(summaries).length} summary documents.`);
}
