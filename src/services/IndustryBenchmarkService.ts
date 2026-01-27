// @ts-nocheck
// Industry Benchmark Service
// Provides industry standard benchmarks for JML process metrics comparison

import { logger } from './LoggingService';

/**
 * Industry benchmark data point
 */
export interface IBenchmarkMetric {
  metricName: string;
  description: string;
  industryAverage: number;
  topQuartile: number; // Best 25%
  bottomQuartile: number; // Worst 25%
  unit: string;
  source: string;
  lastUpdated: Date;
}

/**
 * Benchmark comparison result
 */
export interface IBenchmarkComparison {
  metricName: string;
  organizationValue: number;
  industryAverage: number;
  topQuartile: number;
  performanceRating: 'Excellent' | 'Above Average' | 'Average' | 'Below Average' | 'Poor';
  percentile: number; // 0-100
  gap: number; // Difference from industry average
  gapPercentage: number;
  unit: string;
}

/**
 * Industry sector for more specific benchmarking
 */
export enum IndustrySector {
  Technology = 'Technology',
  Healthcare = 'Healthcare',
  Finance = 'Finance',
  Manufacturing = 'Manufacturing',
  Retail = 'Retail',
  Professional = 'Professional Services',
  General = 'General'
}

/**
 * Company size for relevant comparisons
 */
export enum CompanySize {
  Small = 'Small (1-50)',
  Medium = 'Medium (51-500)',
  Large = 'Large (501-5000)',
  Enterprise = 'Enterprise (5000+)'
}

export class IndustryBenchmarkService {
  private sector: IndustrySector;
  private companySize: CompanySize;

  // Industry benchmark data (based on HR and automation industry research)
  private benchmarks: Map<string, IBenchmarkMetric>;

  constructor(sector: IndustrySector = IndustrySector.General, companySize: CompanySize = CompanySize.Medium) {
    this.sector = sector;
    this.companySize = companySize;
    this.benchmarks = this.initializeBenchmarks();
  }

  /**
   * Initialize industry benchmark data
   */
  private initializeBenchmarks(): Map<string, IBenchmarkMetric> {
    const benchmarks = new Map<string, IBenchmarkMetric>();

    // Time to Onboard (Days)
    benchmarks.set('timeToOnboard', {
      metricName: 'Time to Onboard',
      description: 'Average days from hire date to first-day ready',
      industryAverage: 30,
      topQuartile: 14, // Best performers
      bottomQuartile: 60, // Worst performers
      unit: 'days',
      source: 'SHRM HR Metrics Benchmark Study 2024',
      lastUpdated: new Date('2024-01-01')
    });

    // Employee Onboarding Cost
    benchmarks.set('onboardingCost', {
      metricName: 'Onboarding Cost per Employee',
      description: 'Total cost to onboard one new employee',
      industryAverage: 4000,
      topQuartile: 2500,
      bottomQuartile: 6000,
      unit: 'USD',
      source: 'HR Technology Survey 2024',
      lastUpdated: new Date('2024-01-01')
    });

    // Process Completion Time (Days)
    benchmarks.set('processCycleTime', {
      metricName: 'Process Cycle Time',
      description: 'Average days to complete a JML process',
      industryAverage: 21,
      topQuartile: 10,
      bottomQuartile: 45,
      unit: 'days',
      source: 'Workflow Automation Benchmark Report 2024',
      lastUpdated: new Date('2024-01-01')
    });

    // Task Completion Rate (On-time)
    benchmarks.set('taskCompletionRate', {
      metricName: 'On-Time Task Completion Rate',
      description: 'Percentage of tasks completed by due date',
      industryAverage: 75,
      topQuartile: 90,
      bottomQuartile: 55,
      unit: 'percentage',
      source: 'Project Management Institute 2024',
      lastUpdated: new Date('2024-01-01')
    });

    // Data Entry Error Rate
    benchmarks.set('dataErrorRate', {
      metricName: 'Data Entry Error Rate',
      description: 'Percentage of records with errors',
      industryAverage: 5,
      topQuartile: 1,
      bottomQuartile: 12,
      unit: 'percentage',
      source: 'Data Quality Benchmark Study 2024',
      lastUpdated: new Date('2024-01-01')
    });

    // Compliance Rate
    benchmarks.set('complianceRate', {
      metricName: 'Policy Compliance Rate',
      description: 'Percentage of processes following all policies',
      industryAverage: 85,
      topQuartile: 95,
      bottomQuartile: 70,
      unit: 'percentage',
      source: 'Compliance Benchmark Report 2024',
      lastUpdated: new Date('2024-01-01')
    });

    // Employee Satisfaction (NPS)
    benchmarks.set('employeeSatisfaction', {
      metricName: 'Employee Satisfaction Score',
      description: 'Net Promoter Score for onboarding experience',
      industryAverage: 35,
      topQuartile: 60,
      bottomQuartile: 10,
      unit: 'NPS',
      source: 'Employee Experience Survey 2024',
      lastUpdated: new Date('2024-01-01')
    });

    // Automation Adoption Rate
    benchmarks.set('automationAdoption', {
      metricName: 'Automation Adoption Rate',
      description: 'Percentage of processes using automation',
      industryAverage: 40,
      topQuartile: 70,
      bottomQuartile: 15,
      unit: 'percentage',
      source: 'HR Automation Trends Report 2024',
      lastUpdated: new Date('2024-01-01')
    });

    // Time Saved per Process (Hours)
    benchmarks.set('timeSavedPerProcess', {
      metricName: 'Time Saved per Process',
      description: 'Hours saved through automation per process',
      industryAverage: 3,
      topQuartile: 6,
      bottomQuartile: 1,
      unit: 'hours',
      source: 'Automation ROI Study 2024',
      lastUpdated: new Date('2024-01-01')
    });

    // ROI from Automation
    benchmarks.set('automationROI', {
      metricName: 'Automation ROI',
      description: 'Return on Investment from automation initiatives',
      industryAverage: 200,
      topQuartile: 400,
      bottomQuartile: 100,
      unit: 'percentage',
      source: 'McKinsey Automation Report 2024',
      lastUpdated: new Date('2024-01-01')
    });

    // First-Day Readiness
    benchmarks.set('firstDayReadiness', {
      metricName: 'First-Day Readiness Rate',
      description: 'Percentage of new hires fully ready on day one',
      industryAverage: 70,
      topQuartile: 90,
      bottomQuartile: 45,
      unit: 'percentage',
      source: 'Onboarding Excellence Study 2024',
      lastUpdated: new Date('2024-01-01')
    });

    // IT Provisioning Time (Days)
    benchmarks.set('itProvisioningTime', {
      metricName: 'IT Provisioning Time',
      description: 'Days to provision IT equipment and access',
      industryAverage: 7,
      topQuartile: 3,
      bottomQuartile: 14,
      unit: 'days',
      source: 'IT Operations Benchmark 2024',
      lastUpdated: new Date('2024-01-01')
    });

    return benchmarks;
  }

  /**
   * Get benchmark for a specific metric
   */
  public getBenchmark(metricKey: string): IBenchmarkMetric | undefined {
    return this.benchmarks.get(metricKey);
  }

  /**
   * Get all available benchmarks
   */
  public getAllBenchmarks(): IBenchmarkMetric[] {
    return Array.from(this.benchmarks.values());
  }

  /**
   * Compare organization value against industry benchmark
   */
  public compareToBenchmark(metricKey: string, organizationValue: number): IBenchmarkComparison | null {
    const benchmark = this.benchmarks.get(metricKey);
    if (!benchmark) {
      logger.warn('IndustryBenchmarkService', `Benchmark not found for metric: ${metricKey}`);
      return null;
    }

    // Calculate performance rating and percentile
    const { performanceRating, percentile } = this.calculatePerformance(
      organizationValue,
      benchmark,
      this.isLowerBetter(metricKey)
    );

    // Calculate gap from industry average
    const gap = organizationValue - benchmark.industryAverage;
    const gapPercentage = ((gap / benchmark.industryAverage) * 100);

    return {
      metricName: benchmark.metricName,
      organizationValue,
      industryAverage: benchmark.industryAverage,
      topQuartile: benchmark.topQuartile,
      performanceRating,
      percentile,
      gap,
      gapPercentage,
      unit: benchmark.unit
    };
  }

  /**
   * Batch compare multiple metrics
   */
  public compareMultipleMetrics(metrics: { [key: string]: number }): IBenchmarkComparison[] {
    const comparisons: IBenchmarkComparison[] = [];

    for (const [metricKey, value] of Object.entries(metrics)) {
      const comparison = this.compareToBenchmark(metricKey, value);
      if (comparison) {
        comparisons.push(comparison);
      }
    }

    return comparisons;
  }

  /**
   * Calculate performance rating and percentile
   */
  private calculatePerformance(
    value: number,
    benchmark: IBenchmarkMetric,
    lowerIsBetter: boolean
  ): { performanceRating: IBenchmarkComparison['performanceRating']; percentile: number } {
    const { industryAverage, topQuartile, bottomQuartile } = benchmark;

    let percentile: number;
    let performanceRating: IBenchmarkComparison['performanceRating'];

    if (lowerIsBetter) {
      // For metrics where lower is better (e.g., time, cost, errors)
      if (value <= topQuartile) {
        performanceRating = 'Excellent';
        percentile = 90;
      } else if (value <= industryAverage) {
        performanceRating = 'Above Average';
        percentile = 70;
      } else if (value <= (industryAverage + bottomQuartile) / 2) {
        performanceRating = 'Average';
        percentile = 50;
      } else if (value <= bottomQuartile) {
        performanceRating = 'Below Average';
        percentile = 30;
      } else {
        performanceRating = 'Poor';
        percentile = 10;
      }
    } else {
      // For metrics where higher is better (e.g., satisfaction, compliance, adoption)
      if (value >= topQuartile) {
        performanceRating = 'Excellent';
        percentile = 90;
      } else if (value >= industryAverage) {
        performanceRating = 'Above Average';
        percentile = 70;
      } else if (value >= (industryAverage + bottomQuartile) / 2) {
        performanceRating = 'Average';
        percentile = 50;
      } else if (value >= bottomQuartile) {
        performanceRating = 'Below Average';
        percentile = 30;
      } else {
        performanceRating = 'Poor';
        percentile = 10;
      }
    }

    return { performanceRating, percentile };
  }

  /**
   * Determine if lower values are better for a metric
   */
  private isLowerBetter(metricKey: string): boolean {
    const lowerIsBetterMetrics = [
      'timeToOnboard',
      'onboardingCost',
      'processCycleTime',
      'dataErrorRate',
      'itProvisioningTime',
      'timeSavedPerProcess'
    ];

    return lowerIsBetterMetrics.includes(metricKey);
  }

  /**
   * Get sector-specific adjustment factor
   */
  private getSectorAdjustment(): number {
    // Adjustment factors based on sector complexity
    const adjustments: { [key: string]: number } = {
      [IndustrySector.Technology]: 0.9, // Typically faster
      [IndustrySector.Healthcare]: 1.2, // More compliance requirements
      [IndustrySector.Finance]: 1.15, // Heavy regulation
      [IndustrySector.Manufacturing]: 1.0, // Standard
      [IndustrySector.Retail]: 0.95, // High volume
      [IndustrySector.Professional]: 1.0, // Standard
      [IndustrySector.General]: 1.0 // Baseline
    };

    return adjustments[this.sector] || 1.0;
  }

  /**
   * Get company size adjustment factor
   */
  private getCompanySizeAdjustment(): number {
    // Smaller companies typically faster but less automated
    const adjustments: { [key: string]: number } = {
      [CompanySize.Small]: 0.8,
      [CompanySize.Medium]: 1.0,
      [CompanySize.Large]: 1.1,
      [CompanySize.Enterprise]: 1.2
    };

    return adjustments[this.companySize] || 1.0;
  }

  /**
   * Get adjusted benchmark (considering sector and size)
   */
  public getAdjustedBenchmark(metricKey: string): IBenchmarkMetric | undefined {
    const benchmark = this.benchmarks.get(metricKey);
    if (!benchmark) return undefined;

    const sectorAdjustment = this.getSectorAdjustment();
    const sizeAdjustment = this.getCompanySizeAdjustment();
    const combinedAdjustment = sectorAdjustment * sizeAdjustment;

    return {
      ...benchmark,
      industryAverage: benchmark.industryAverage * combinedAdjustment,
      topQuartile: benchmark.topQuartile * combinedAdjustment,
      bottomQuartile: benchmark.bottomQuartile * combinedAdjustment
    };
  }

  /**
   * Get overall performance summary
   */
  public getPerformanceSummary(comparisons: IBenchmarkComparison[]): {
    overallRating: string;
    excellentCount: number;
    aboveAverageCount: number;
    averageCount: number;
    belowAverageCount: number;
    poorCount: number;
    averagePercentile: number;
  } {
    const ratings = {
      excellentCount: 0,
      aboveAverageCount: 0,
      averageCount: 0,
      belowAverageCount: 0,
      poorCount: 0
    };

    let totalPercentile = 0;

    for (const comparison of comparisons) {
      totalPercentile += comparison.percentile;

      switch (comparison.performanceRating) {
        case 'Excellent':
          ratings.excellentCount++;
          break;
        case 'Above Average':
          ratings.aboveAverageCount++;
          break;
        case 'Average':
          ratings.averageCount++;
          break;
        case 'Below Average':
          ratings.belowAverageCount++;
          break;
        case 'Poor':
          ratings.poorCount++;
          break;
      }
    }

    const averagePercentile = comparisons.length > 0 ? totalPercentile / comparisons.length : 0;

    let overallRating: string;
    if (averagePercentile >= 80) {
      overallRating = 'Excellent';
    } else if (averagePercentile >= 65) {
      overallRating = 'Above Average';
    } else if (averagePercentile >= 45) {
      overallRating = 'Average';
    } else if (averagePercentile >= 25) {
      overallRating = 'Below Average';
    } else {
      overallRating = 'Needs Improvement';
    }

    return {
      overallRating,
      ...ratings,
      averagePercentile
    };
  }
}
