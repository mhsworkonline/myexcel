/** Registers Phase 3 features (charts, pivots, what-if, printing). */
import { ChartDataDialog, ChartFormatDialog, InsertChartDialog } from './charts/chartDialogs';
import { registerDialogs } from './dialogs/DialogHost';
import {
  DataTableDialog,
  GoalSeekDialog,
  GoalSeekStatusDialog,
  ScenarioEditDialog,
  ScenarioManagerDialog,
  ScenarioSummaryDialog,
  SolverDialog,
  SolverResultsDialog,
  SparklineDialog,
} from './dialogs/whatif';
import { CreatePivotDialog, PivotCalcFieldDialog, PivotFilterDialog, PivotGroupDialog, PivotValueSettingsDialog } from './pivot/PivotPanel';
import { PageSetupDialog, PrintView } from './PrintView';

let installed = false;

export function installPhase3(): void {
  if (installed) return;
  installed = true;
  registerDialogs({
    insertChart: InsertChartDialog,
    chartData: ChartDataDialog,
    chartFormat: ChartFormatDialog,
    createPivot: CreatePivotDialog,
    pivotValueSettings: PivotValueSettingsDialog,
    pivotGroup: PivotGroupDialog,
    pivotCalcField: PivotCalcFieldDialog,
    pivotFilter: PivotFilterDialog,
    goalSeek: GoalSeekDialog,
    goalSeekStatus: GoalSeekStatusDialog,
    solver: SolverDialog,
    solverResults: SolverResultsDialog,
    scenarioManager: ScenarioManagerDialog,
    scenarioEdit: ScenarioEditDialog,
    scenarioSummary: ScenarioSummaryDialog,
    dataTable: DataTableDialog,
    sparklines: SparklineDialog,
    print: () => <PrintView />,
    pageSetup: PageSetupDialog,
  });
}
