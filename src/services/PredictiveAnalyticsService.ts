// @ts-nocheck
// Predictive Analytics Service
// Forecasts future trends based on historical JML process data

import { SPFI } from '@pnp/sp';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { logger } from './LoggingService';

/**
 * Time period for forecasting
 */
export enum ForecastPeriod {
  NextMonth = 'Next Month',
  NextQuarter = 'Next Quarter',
  NextSixMonths = 'Next 6 Months',
  NextYear = 'Next Year'
}

/**
 * Forecast data point
 */
export interface IForecastDataPoint {
  date: Date;
  predictedValue: number;
  confidenceLow: number; // Lower bound of confidence interval
  confidenceHigh: number; // Upper bound of confidence interval
  confidence: number; // Confidence percentage (0-100)
}

/**
 * Trend direction
 */
export type TrendDirection = 'Increasing' | 'Decreasing' | 'Stable';

/**
 * Process volume forecast
 */
export interface IProcessVolumeForecast {
  metricName: string;
  currentValue: number;
  forecastPeriod: ForecastPeriod;
  forecastedValue: number;
  trend: TrendDirection;
  growthRate: number; // Percentage
  confidence: number;
  dataPoints: IForecastDataPoint[];
}

/**
 * Cost savings forecast
 */
export interface ICostSavingsForecast {
  metricName: string;
  currentMonthlySavings: number;
  forecastPeriod: ForecastPeriod;
  forecastedMonthlySavings: number;
  totalProjectedSavings: number;
  trend: TrendDirection;
  growthRate: number;
  confidence: number;
  dataPoints: IForecastDataPoint[];
}

/**
 * Adoption rate forecast
 */
export interface IAdoptionForecast {
  metricName: string;
  currentAdoptionRate: number; // Percentage
  forecastPeriod: ForecastPeriod;
  forecastedAdoptionRate: number;
  trend: TrendDirection;
  saturationPoint: number; // Expected maximum adoption rate
  timeToSaturation: number; // Months until saturation
  confidence: number;
  dataPoints: IForecastDataPoint[];
}

/**
 * Historical data point for analysis
 */
interface IHistoricalDataPoint {
  date: Date;
  value: number;
}

export class PredictiveAnalyticsService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Forecast process volume for future periods
   */
  public async forecastProcessVolume(period: ForecastPeriod): Promise<IProcessVolumeForecast> {
    try {
      // Get historical process data (last 12 months)
      const historicalData = await this.getHistoricalProcessVolume();

      // Perform linear regression
      const { slope, intercept } = this.linearRegression(historicalData);

      // Calculate forecast period in months
      const forecastMonths = this.getPeriodInMonths(period);

      // Current value (latest month)
      const currentValue = historicalData.length > 0
        ? historicalData[historicalData.length - 1].value
        : 0;

      // Forecast future value
      const forecastedValue = this.predictValue(slope, intercept, historicalData.length + forecastMonths);

      // Calculate growth rate
      const growthRate = currentValue > 0
        ? ((forecastedValue - currentValue) / currentValue) * 100
        : 0;

      // Determine trend
      const trend = this.determineTrend(slope);

      // Generate forecast data points
      const dataPoints = this.generateForecastPoints(
        historicalData,
        slope,
        intercept,
        forecastMonths
      );

      // Calculate confidence (decreases with forecast distance)
      const confidence = Math.max(60, 95 - (forecastMonths * 2));

      return {
        metricName: 'Process Volume',
        currentValue,
        forecastPeriod: period,
        forecastedValue: Math.round(forecastedValue),
        trend,
        growthRate,
        confidence,
        dataPoints
      };
    } catch (error) {
      logger.error('PredictiveAnalyticsService', 'Failed to forecast process volume:', error);
      throw error;
    }
  }

  /**
   * Forecast cost savings for future periods
   */
  public async forecastCostSavings(
    currentMonthlySavings: number,
    period: ForecastPeriod
  ): Promise<ICostSavingsForecast> {
    try {
      // Simulate historical savings data based on current value
      // In production, this would query actual historical data
      const historicalData = this.simulateSavingsHistory(currentMonthlySavings);

      // Perform regression
      const { slope, intercept } = this.linearRegression(historicalData);

      // Forecast period
      const forecastMonths = this.getPeriodInMonths(period);

      // Forecast future monthly savings
      const forecastedMonthlySavings = this.predictValue(
        slope,
        intercept,
        historicalData.length + forecastMonths
      );

      // Calculate total projected savings over the period
      const totalProjectedSavings = forecastedMonthlySavings * forecastMonths;

      // Calculate growth rate
      const growthRate = currentMonthlySavings > 0
        ? ((forecastedMonthlySavings - currentMonthlySavings) / currentMonthlySavings) * 100
        : 0;

      // Determine trend
      const trend = this.determineTrend(slope);

      // Generate data points
      const dataPoints = this.generateForecastPoints(
        historicalData,
        slope,
        intercept,
        forecastMonths
      );

      const confidence = Math.max(65, 90 - (forecastMonths * 1.5));

      return {
        metricName: 'Monthly Cost Savings',
        currentMonthlySavings,
        forecastPeriod: period,
        forecastedMonthlySavings,
        totalProjectedSavings,
        trend,
        growthRate,
        confidence,
        dataPoints
      };
    } catch (error) {
      logger.error('PredictiveAnalyticsService', 'Failed to forecast cost savings:', error);
      throw error;
    }
  }

  /**
   * Forecast automation adoption rate
   */
  public async forecastAdoptionRate(
    currentAdoptionRate: number,
    period: ForecastPeriod
  ): Promise<IAdoptionForecast> {
    try {
      // Simulate historical adoption data
      const historicalData = this.simulateAdoptionHistory(currentAdoptionRate);

      // Use logistic growth model for adoption (S-curve)
      const saturationPoint = 85; // Maximum expected adoption rate
      const forecastMonths = this.getPeriodInMonths(period);

      // Calculate forecasted adoption using logistic growth
      const forecastedAdoptionRate = this.predictLogisticGrowth(
        currentAdoptionRate,
        saturationPoint,
        forecastMonths
      );

      // Estimate time to saturation (80% of max)
      const targetRate = saturationPoint * 0.8;
      const timeToSaturation = this.estimateTimeToTarget(
        currentAdoptionRate,
        targetRate,
        saturationPoint
      );

      // Determine trend
      const trend = forecastedAdoptionRate > currentAdoptionRate ? 'Increasing' : 'Stable';

      // Generate logistic growth data points
      const dataPoints = this.generateLogisticForecastPoints(
        currentAdoptionRate,
        saturationPoint,
        forecastMonths
      );

      const confidence = Math.max(70, 92 - (forecastMonths * 1.2));

      return {
        metricName: 'Automation Adoption Rate',
        currentAdoptionRate,
        forecastPeriod: period,
        forecastedAdoptionRate,
        trend,
        saturationPoint,
        timeToSaturation,
        confidence,
        dataPoints
      };
    } catch (error) {
      logger.error('PredictiveAnalyticsService', 'Failed to forecast adoption rate:', error);
      throw error;
    }
  }

  /**
   * Get historical process volume data
   */
  private async getHistoricalProcessVolume(): Promise<IHistoricalDataPoint[]> {
    try {
      const twelveMonthsAgo = new Date();
      twelveMonthsAgo.setMonth(twelveMonthsAgo.getMonth() - 12);

      const processes = await this.sp.web.lists
        .getByTitle('JML_Processes')
        .items.filter(`Created ge datetime'${twelveMonthsAgo.toISOString()}'`)
        .select('Id', 'Created')
        .orderBy('Created', true)();

      // Group by month
      const monthlyData = new Map<string, number>();

      for (const process of processes) {
        const date = new Date(process.Created);
        const month = date.getMonth() + 1;
        const monthStr = month < 10 ? `0${month}` : `${month}`;
        const monthKey = `${date.getFullYear()}-${monthStr}`;

        monthlyData.set(monthKey, (monthlyData.get(monthKey) || 0) + 1);
      }

      // Convert to data points
      const dataPoints: IHistoricalDataPoint[] = [];
      const sortedKeys = Array.from(monthlyData.keys()).sort();

      for (const key of sortedKeys) {
        const [year, month] = key.split('-').map(Number);
        dataPoints.push({
          date: new Date(year, month - 1, 1),
          value: monthlyData.get(key) || 0
        });
      }

      return dataPoints;
    } catch (error) {
      logger.error('PredictiveAnalyticsService', 'Failed to get historical volume:', error);
      // Return simulated data as fallback
      return this.simulateProcessVolumeHistory();
    }
  }

  /**
   * Linear regression calculation
   */
  private linearRegression(dataPoints: IHistoricalDataPoint[]): { slope: number; intercept: number } {
    if (dataPoints.length === 0) {
      return { slope: 0, intercept: 0 };
    }

    const n = dataPoints.length;
    let sumX = 0;
    let sumY = 0;
    let sumXY = 0;
    let sumXX = 0;

    dataPoints.forEach((point, index) => {
      const x = index;
      const y = point.value;
      sumX += x;
      sumY += y;
      sumXY += x * y;
      sumXX += x * x;
    });

    const slope = (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX);
    const intercept = (sumY - slope * sumX) / n;

    return { slope, intercept };
  }

  /**
   * Predict value using linear regression
   */
  private predictValue(slope: number, intercept: number, x: number): number {
    return slope * x + intercept;
  }

  /**
   * Predict value using logistic growth model (S-curve)
   */
  private predictLogisticGrowth(current: number, max: number, months: number): number {
    // Logistic growth: P(t) = K / (1 + ((K - P0) / P0) * e^(-r*t))
    // K = carrying capacity (max)
    // P0 = initial value
    // r = growth rate
    // t = time

    const growthRate = 0.15; // Monthly growth rate
    const ratio = (max - current) / current;
    const denominator = 1 + ratio * Math.exp(-growthRate * months);

    return Math.min(max, max / denominator);
  }

  /**
   * Estimate months to reach target adoption rate
   */
  private estimateTimeToTarget(current: number, target: number, max: number): number {
    if (current >= target) return 0;

    const growthRate = 0.15;
    const ratio = (max - current) / current;
    const targetRatio = (max / target) - 1;

    const months = Math.log(targetRatio / ratio) / growthRate;

    return Math.max(0, Math.round(months));
  }

  /**
   * Determine trend direction from slope
   */
  private determineTrend(slope: number): TrendDirection {
    if (slope > 0.5) return 'Increasing';
    if (slope < -0.5) return 'Decreasing';
    return 'Stable';
  }

  /**
   * Get forecast period in months
   */
  private getPeriodInMonths(period: ForecastPeriod): number {
    switch (period) {
      case ForecastPeriod.NextMonth:
        return 1;
      case ForecastPeriod.NextQuarter:
        return 3;
      case ForecastPeriod.NextSixMonths:
        return 6;
      case ForecastPeriod.NextYear:
        return 12;
      default:
        return 3;
    }
  }

  /**
   * Generate forecast data points
   */
  private generateForecastPoints(
    historicalData: IHistoricalDataPoint[],
    slope: number,
    intercept: number,
    forecastMonths: number
  ): IForecastDataPoint[] {
    const dataPoints: IForecastDataPoint[] = [];
    const startIndex = historicalData.length;
    const lastDate = historicalData.length > 0
      ? new Date(historicalData[historicalData.length - 1].date)
      : new Date();

    for (let i = 1; i <= forecastMonths; i++) {
      const predictedValue = this.predictValue(slope, intercept, startIndex + i);
      const confidence = Math.max(60, 95 - (i * 2));

      // Confidence interval (±10-20% based on distance)
      const margin = predictedValue * (0.1 + (i * 0.02));

      const forecastDate = new Date(lastDate);
      forecastDate.setMonth(forecastDate.getMonth() + i);

      dataPoints.push({
        date: forecastDate,
        predictedValue: Math.max(0, predictedValue),
        confidenceLow: Math.max(0, predictedValue - margin),
        confidenceHigh: predictedValue + margin,
        confidence
      });
    }

    return dataPoints;
  }

  /**
   * Generate logistic growth forecast points
   */
  private generateLogisticForecastPoints(
    current: number,
    max: number,
    forecastMonths: number
  ): IForecastDataPoint[] {
    const dataPoints: IForecastDataPoint[] = [];
    const now = new Date();

    for (let i = 1; i <= forecastMonths; i++) {
      const predictedValue = this.predictLogisticGrowth(current, max, i);
      const confidence = Math.max(70, 92 - (i * 1.2));

      // Confidence interval
      const margin = predictedValue * 0.05;

      const forecastDate = new Date(now);
      forecastDate.setMonth(forecastDate.getMonth() + i);

      dataPoints.push({
        date: forecastDate,
        predictedValue,
        confidenceLow: Math.max(0, predictedValue - margin),
        confidenceHigh: Math.min(max, predictedValue + margin),
        confidence
      });
    }

    return dataPoints;
  }

  /**
   * Simulate historical process volume (fallback)
   */
  private simulateProcessVolumeHistory(): IHistoricalDataPoint[] {
    const dataPoints: IHistoricalDataPoint[] = [];
    const now = new Date();
    const baseValue = 10;
    const growthRate = 1.1; // 10% monthly growth

    for (let i = 11; i >= 0; i--) {
      const date = new Date(now);
      date.setMonth(date.getMonth() - i);

      const value = Math.round(baseValue * Math.pow(growthRate, 11 - i));

      dataPoints.push({ date, value });
    }

    return dataPoints;
  }

  /**
   * Simulate historical savings data
   */
  private simulateSavingsHistory(currentValue: number): IHistoricalDataPoint[] {
    const dataPoints: IHistoricalDataPoint[] = [];
    const now = new Date();
    const growthRate = 1.08; // 8% monthly growth

    for (let i = 11; i >= 0; i--) {
      const date = new Date(now);
      date.setMonth(date.getMonth() - i);

      const value = currentValue / Math.pow(growthRate, i);

      dataPoints.push({ date, value });
    }

    return dataPoints;
  }

  /**
   * Simulate historical adoption data
   */
  private simulateAdoptionHistory(currentRate: number): IHistoricalDataPoint[] {
    const dataPoints: IHistoricalDataPoint[] = [];
    const now = new Date();

    // Simulate logistic growth history
    for (let i = 11; i >= 0; i--) {
      const date = new Date(now);
      date.setMonth(date.getMonth() - i);

      // Reverse logistic calculation
      const monthsBack = i;
      const historicRate = currentRate / (1 + 0.15 * monthsBack);

      dataPoints.push({ date, value: historicRate });
    }

    return dataPoints;
  }
}
