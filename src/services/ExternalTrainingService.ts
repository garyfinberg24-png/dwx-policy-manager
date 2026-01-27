// @ts-nocheck
/**
 * ExternalTrainingService - Service for managing external training content
 *
 * This service provides:
 * - Access to curated free courses from major platforms
 * - Filtering and searching capabilities
 * - Platform-specific metadata and branding
 * - Future integration points for API-based content fetching
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { getSP } from '../utils/pnpjsConfig';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

// ============================================
// Interfaces
// ============================================

export interface IExternalProvider {
  id: string;
  name: string;
  description: string;
  logoUrl: string;
  websiteUrl: string;
  primaryColor: string;
  isFree: boolean;
  hasApi: boolean;
  apiEndpoint?: string;
}

export interface IExternalCourse {
  id: number;
  title: string;
  description: string;
  courseCode: string;
  courseType: string;
  contentFormat: string;
  difficultyLevel: string;
  duration: number; // in minutes
  provider: string;
  instructor?: string;
  thumbnailUrl?: string;
  contentUrl?: string;
  language: string;
  isMandatory: boolean;
  isActive: boolean;
  isFree: boolean;
  points: number;
  xpReward: number;
  passingScore: number;
  tags: string[];
  rating?: number;
  enrollments?: number;
  lastUpdated?: Date;
}

export interface ICourseFilter {
  providers?: string[];
  difficultyLevels?: string[];
  contentFormats?: string[];
  tags?: string[];
  isFree?: boolean;
  minDuration?: number;
  maxDuration?: number;
  searchQuery?: string;
}

export interface ICourseCategory {
  id: string;
  name: string;
  description: string;
  icon: string;
  color: string;
  courseCount: number;
  tags: string[];
}

// ============================================
// External Provider Definitions
// ============================================

export const ExternalProviders: IExternalProvider[] = [
  {
    id: 'microsoft-learn',
    name: 'Microsoft Learn',
    description: 'Free, interactive, hands-on training from Microsoft covering Azure, Microsoft 365, Power Platform, and more.',
    logoUrl: 'https://docs.microsoft.com/en-us/media/logos/logo-ms-social.png',
    websiteUrl: 'https://learn.microsoft.com',
    primaryColor: '#0078D4',
    isFree: true,
    hasApi: true,
    apiEndpoint: 'https://learn.microsoft.com/api/catalog/'
  },
  {
    id: 'linkedin-learning',
    name: 'LinkedIn Learning',
    description: 'Professional development courses on business, technology, and creative skills.',
    logoUrl: 'https://content.linkedin.com/content/dam/me/business/en-us/amp/brand-site/v2/bg/LI-Bug.svg.original.svg',
    websiteUrl: 'https://www.linkedin.com/learning',
    primaryColor: '#0A66C2',
    isFree: false, // Most content requires subscription
    hasApi: true
  },
  {
    id: 'coursera',
    name: 'Coursera',
    description: 'Online courses from top universities and companies worldwide.',
    logoUrl: 'https://d3njjcbhbojbot.cloudfront.net/api/utilities/v1/imageproxy/https://coursera.s3.amazonaws.com/media/coursera-logo-square.png',
    websiteUrl: 'https://www.coursera.org',
    primaryColor: '#0056D2',
    isFree: false, // Audit is free, certificates paid
    hasApi: true,
    apiEndpoint: 'https://api.coursera.org/api/'
  },
  {
    id: 'freecodecamp',
    name: 'freeCodeCamp',
    description: 'Free coding bootcamp with interactive lessons and certification projects.',
    logoUrl: 'https://design-style-guide.freecodecamp.org/downloads/fcc_secondary_small.svg',
    websiteUrl: 'https://www.freecodecamp.org',
    primaryColor: '#0A0A23',
    isFree: true,
    hasApi: false
  },
  {
    id: 'edx',
    name: 'edX',
    description: 'Online courses from Harvard, MIT, and other leading institutions.',
    logoUrl: 'https://www.edx.org/images/logos/edx-logo-elm.svg',
    websiteUrl: 'https://www.edx.org',
    primaryColor: '#02262B',
    isFree: false, // Audit is free
    hasApi: true,
    apiEndpoint: 'https://courses.edx.org/api/'
  },
  {
    id: 'aws-skillbuilder',
    name: 'AWS Skill Builder',
    description: 'Free and paid AWS training courses and learning paths.',
    logoUrl: 'https://d1.awsstatic.com/logos/aws-logo-lockups/poweredbyaws/PB_AWS_logo_RGB_stacked.547f032d90171cdea4dd90c258f47373c5573db5.png',
    websiteUrl: 'https://explore.skillbuilder.aws',
    primaryColor: '#FF9900',
    isFree: true, // Free tier available
    hasApi: false
  },
  {
    id: 'google-digital-garage',
    name: 'Google Digital Garage',
    description: 'Free online courses on digital marketing, data, and career development from Google.',
    logoUrl: 'https://www.gstatic.com/images/branding/product/2x/google_g_64dp.png',
    websiteUrl: 'https://learndigital.withgoogle.com/digitalgarage',
    primaryColor: '#4285F4',
    isFree: true,
    hasApi: false
  },
  {
    id: 'cisco-netacad',
    name: 'Cisco Networking Academy',
    description: 'Free networking, cybersecurity, and IoT courses from Cisco.',
    logoUrl: 'https://www.netacad.com/sites/default/files/images/cisco_netacad_logo.png',
    websiteUrl: 'https://www.netacad.com',
    primaryColor: '#049FD9',
    isFree: true,
    hasApi: false
  },
  {
    id: 'khan-academy',
    name: 'Khan Academy',
    description: 'Free world-class education in math, science, computing, and more.',
    logoUrl: 'https://cdn.kastatic.org/images/khan-logo-dark-background.new.png',
    websiteUrl: 'https://www.khanacademy.org',
    primaryColor: '#14BF96',
    isFree: true,
    hasApi: true,
    apiEndpoint: 'https://www.khanacademy.org/api/v1/'
  },
  {
    id: 'pluralsight',
    name: 'Pluralsight',
    description: 'Technology skills platform for software development, IT, and creative professionals.',
    logoUrl: 'https://www.pluralsight.com/etc/clientlibs/pluralsight/main/images/global/header/PS_logo.png',
    websiteUrl: 'https://www.pluralsight.com',
    primaryColor: '#F15B2A',
    isFree: false,
    hasApi: true
  },
  {
    id: 'udemy',
    name: 'Udemy',
    description: 'Online learning marketplace with courses on every topic.',
    logoUrl: 'https://www.udemy.com/staticx/udemy/images/v7/logo-udemy.svg',
    websiteUrl: 'https://www.udemy.com',
    primaryColor: '#A435F0',
    isFree: false, // Some free courses
    hasApi: true,
    apiEndpoint: 'https://www.udemy.com/api-2.0/'
  }
];

// ============================================
// Course Categories
// ============================================

export const CourseCategories: ICourseCategory[] = [
  {
    id: 'cloud',
    name: 'Cloud Computing',
    description: 'AWS, Azure, Google Cloud certifications and training',
    icon: 'Cloud',
    color: '#0078D4',
    courseCount: 0,
    tags: ['Azure', 'AWS', 'Cloud', 'Google Cloud', 'DevOps']
  },
  {
    id: 'programming',
    name: 'Programming & Development',
    description: 'Learn to code with Python, JavaScript, and more',
    icon: 'Code',
    color: '#107C10',
    courseCount: 0,
    tags: ['Python', 'JavaScript', 'Programming', 'Web Development', 'SQL']
  },
  {
    id: 'data-science',
    name: 'Data Science & AI',
    description: 'Machine learning, data analysis, and artificial intelligence',
    icon: 'Processing',
    color: '#8764B8',
    courseCount: 0,
    tags: ['Machine Learning', 'AI', 'Data Science', 'Data Analysis', 'Statistics']
  },
  {
    id: 'cybersecurity',
    name: 'Cybersecurity',
    description: 'Security fundamentals, compliance, and certifications',
    icon: 'Shield',
    color: '#D83B01',
    courseCount: 0,
    tags: ['Security', 'Cybersecurity', 'Compliance', 'Identity']
  },
  {
    id: 'microsoft-365',
    name: 'Microsoft 365 & Power Platform',
    description: 'Productivity tools, Power Apps, Power Automate',
    icon: 'Waffle',
    color: '#0078D4',
    courseCount: 0,
    tags: ['Microsoft 365', 'Power Platform', 'Power Apps', 'Power Automate', 'SharePoint']
  },
  {
    id: 'project-management',
    name: 'Project Management',
    description: 'Agile, Scrum, PMP, and project management skills',
    icon: 'ProjectManagement',
    color: '#FFB900',
    courseCount: 0,
    tags: ['Project Management', 'Agile', 'Scrum', 'Certificate']
  },
  {
    id: 'soft-skills',
    name: 'Soft Skills & Leadership',
    description: 'Communication, leadership, and professional development',
    icon: 'People',
    color: '#00B7C3',
    courseCount: 0,
    tags: ['Communication', 'Leadership', 'Soft Skills', 'Time Management']
  },
  {
    id: 'digital-marketing',
    name: 'Digital Marketing',
    description: 'SEO, analytics, social media, and marketing skills',
    icon: 'Megaphone',
    color: '#E3008C',
    courseCount: 0,
    tags: ['Digital Marketing', 'SEO', 'SEM', 'Google Analytics', 'Marketing']
  }
];

// ============================================
// Service Class
// ============================================

export class ExternalTrainingService {
  private _sp: SPFI;
  private _context: WebPartContext;

  constructor(context: WebPartContext) {
    this._context = context;
    this._sp = getSP(context);
  }

  /**
   * Get all external training providers
   */
  public getProviders(): IExternalProvider[] {
    return ExternalProviders;
  }

  /**
   * Get provider by ID
   */
  public getProviderById(id: string): IExternalProvider | undefined {
    return ExternalProviders.find(p => p.id === id);
  }

  /**
   * Get free-only providers
   */
  public getFreeProviders(): IExternalProvider[] {
    return ExternalProviders.filter(p => p.isFree);
  }

  /**
   * Get all course categories
   */
  public getCategories(): ICourseCategory[] {
    return CourseCategories;
  }

  /**
   * Get all external/free courses from SharePoint
   */
  public async getCourses(filter?: ICourseFilter): Promise<IExternalCourse[]> {
    try {
      let filterQuery = "IsActive eq 1";

      // Add provider filter
      if (filter?.providers && filter.providers.length > 0) {
        const providerFilters = filter.providers.map(p => `Provider eq '${p}'`).join(' or ');
        filterQuery += ` and (${providerFilters})`;
      }

      // Add difficulty filter
      if (filter?.difficultyLevels && filter.difficultyLevels.length > 0) {
        const diffFilters = filter.difficultyLevels.map(d => `DifficultyLevel eq '${d}'`).join(' or ');
        filterQuery += ` and (${diffFilters})`;
      }

      // Add content format filter
      if (filter?.contentFormats && filter.contentFormats.length > 0) {
        const formatFilters = filter.contentFormats.map(f => `ContentFormat eq '${f}'`).join(' or ');
        filterQuery += ` and (${formatFilters})`;
      }

      // Duration filters
      if (filter?.minDuration) {
        filterQuery += ` and Duration ge ${filter.minDuration}`;
      }
      if (filter?.maxDuration) {
        filterQuery += ` and Duration le ${filter.maxDuration}`;
      }

      const items = await this._sp.web.lists
        .getByTitle('JML_TrainingCourses')
        .items
        .filter(filterQuery)
        .select(
          'Id', 'Title', 'Description', 'CourseCode', 'CourseType', 'ContentFormat',
          'DifficultyLevel', 'Duration', 'Provider', 'Instructor', 'ThumbnailUrl',
          'ContentUrl', 'Language', 'IsMandatory', 'IsActive', 'Points', 'XPReward',
          'PassingScore', 'Tags', 'AverageRating', 'TotalEnrollments', 'Modified'
        )
        .orderBy('Title')
        .top(500)();

      let courses = items.map(item => this.mapToCourse(item));

      // Apply tag filter (client-side since Tags is a text field)
      if (filter?.tags && filter.tags.length > 0) {
        courses = courses.filter(c =>
          filter.tags!.some(tag => c.tags.includes(tag))
        );
      }

      // Apply search filter (client-side)
      if (filter?.searchQuery) {
        const query = filter.searchQuery.toLowerCase();
        courses = courses.filter(c =>
          c.title.toLowerCase().includes(query) ||
          c.description.toLowerCase().includes(query) ||
          c.provider.toLowerCase().includes(query) ||
          c.tags.some(t => t.toLowerCase().includes(query))
        );
      }

      // Filter free courses
      if (filter?.isFree) {
        courses = courses.filter(c => c.isFree);
      }

      return courses;
    } catch (error) {
      console.error('Error fetching external courses:', error);
      throw error;
    }
  }

  /**
   * Get courses by provider
   */
  public async getCoursesByProvider(provider: string): Promise<IExternalCourse[]> {
    return this.getCourses({ providers: [provider] });
  }

  /**
   * Get courses by category
   */
  public async getCoursesByCategory(categoryId: string): Promise<IExternalCourse[]> {
    const category = CourseCategories.find(c => c.id === categoryId);
    if (!category) return [];

    return this.getCourses({ tags: category.tags });
  }

  /**
   * Get free courses only
   */
  public async getFreeCourses(): Promise<IExternalCourse[]> {
    // Free providers
    const freeProviderNames = this.getFreeProviders().map(p => p.name);
    return this.getCourses({ providers: freeProviderNames });
  }

  /**
   * Get course by ID
   */
  public async getCourseById(id: number): Promise<IExternalCourse | null> {
    try {
      const item = await this._sp.web.lists
        .getByTitle('JML_TrainingCourses')
        .items
        .getById(id)
        .select(
          'Id', 'Title', 'Description', 'CourseCode', 'CourseType', 'ContentFormat',
          'DifficultyLevel', 'Duration', 'Provider', 'Instructor', 'ThumbnailUrl',
          'ContentUrl', 'Language', 'IsMandatory', 'IsActive', 'Points', 'XPReward',
          'PassingScore', 'Tags', 'AverageRating', 'TotalEnrollments', 'Modified'
        )();

      return this.mapToCourse(item);
    } catch (error) {
      console.error('Error fetching course by ID:', error);
      return null;
    }
  }

  /**
   * Search courses
   */
  public async searchCourses(query: string): Promise<IExternalCourse[]> {
    return this.getCourses({ searchQuery: query });
  }

  /**
   * Get popular courses (by enrollment count)
   */
  public async getPopularCourses(limit: number = 10): Promise<IExternalCourse[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle('JML_TrainingCourses')
        .items
        .filter("IsActive eq 1")
        .select(
          'Id', 'Title', 'Description', 'CourseCode', 'CourseType', 'ContentFormat',
          'DifficultyLevel', 'Duration', 'Provider', 'Instructor', 'ThumbnailUrl',
          'ContentUrl', 'Language', 'IsMandatory', 'IsActive', 'Points', 'XPReward',
          'PassingScore', 'Tags', 'AverageRating', 'TotalEnrollments', 'Modified'
        )
        .orderBy('TotalEnrollments', false)
        .top(limit)();

      return items.map(item => this.mapToCourse(item));
    } catch (error) {
      console.error('Error fetching popular courses:', error);
      return [];
    }
  }

  /**
   * Get recommended courses based on user's skills and interests
   */
  public async getRecommendedCourses(
    userSkills: string[],
    completedCourseIds: number[],
    limit: number = 10
  ): Promise<IExternalCourse[]> {
    try {
      const allCourses = await this.getCourses();

      // Filter out completed courses
      let available = allCourses.filter(c => !completedCourseIds.includes(c.id));

      // Score courses based on skill match
      const scored = available.map(course => {
        let score = 0;

        // Match tags with user skills
        userSkills.forEach(skill => {
          if (course.tags.some(t => t.toLowerCase().includes(skill.toLowerCase()))) {
            score += 10;
          }
        });

        // Boost free courses
        if (course.isFree) score += 5;

        // Boost highly rated courses
        if (course.rating && course.rating >= 4.5) score += 3;

        // Boost popular courses
        if (course.enrollments && course.enrollments > 1000) score += 2;

        return { course, score };
      });

      // Sort by score and return top results
      return scored
        .sort((a, b) => b.score - a.score)
        .slice(0, limit)
        .map(s => s.course);
    } catch (error) {
      console.error('Error getting recommended courses:', error);
      return [];
    }
  }

  /**
   * Get course statistics by provider
   */
  public async getCourseStatsByProvider(): Promise<{ provider: string; count: number; totalDuration: number }[]> {
    try {
      const courses = await this.getCourses();
      const stats: { [key: string]: { count: number; totalDuration: number } } = {};

      courses.forEach(course => {
        if (!stats[course.provider]) {
          stats[course.provider] = { count: 0, totalDuration: 0 };
        }
        stats[course.provider].count++;
        stats[course.provider].totalDuration += course.duration;
      });

      return Object.entries(stats).map(([provider, data]) => ({
        provider,
        count: data.count,
        totalDuration: data.totalDuration
      }));
    } catch (error) {
      console.error('Error getting course stats:', error);
      return [];
    }
  }

  /**
   * Get all unique tags from courses
   */
  public async getAvailableTags(): Promise<string[]> {
    try {
      const courses = await this.getCourses();
      const tagSet = new Set<string>();

      courses.forEach(course => {
        course.tags.forEach(tag => tagSet.add(tag));
      });

      return Array.from(tagSet).sort();
    } catch (error) {
      console.error('Error getting available tags:', error);
      return [];
    }
  }

  /**
   * Launch external course (open in new tab)
   */
  public launchCourse(course: IExternalCourse): void {
    if (course.contentUrl) {
      window.open(course.contentUrl, '_blank', 'noopener,noreferrer');
    }
  }

  /**
   * Map SharePoint item to IExternalCourse
   */
  private mapToCourse(item: any): IExternalCourse {
    const freeProviders = this.getFreeProviders().map(p => p.name);

    return {
      id: item.Id,
      title: item.Title || '',
      description: item.Description || '',
      courseCode: item.CourseCode || '',
      courseType: item.CourseType || 'eLearning',
      contentFormat: item.ContentFormat || 'Video',
      difficultyLevel: item.DifficultyLevel || 'Beginner',
      duration: item.Duration || 0,
      provider: item.Provider || '',
      instructor: item.Instructor,
      thumbnailUrl: item.ThumbnailUrl,
      contentUrl: item.ContentUrl,
      language: item.Language || 'English',
      isMandatory: item.IsMandatory || false,
      isActive: item.IsActive !== false,
      isFree: freeProviders.includes(item.Provider) || (item.Tags || '').toLowerCase().includes('free'),
      points: item.Points || 0,
      xpReward: item.XPReward || 0,
      passingScore: item.PassingScore || 0,
      tags: item.Tags ? item.Tags.split(';').map((t: string) => t.trim()).filter((t: string) => t) : [],
      rating: item.AverageRating,
      enrollments: item.TotalEnrollments,
      lastUpdated: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  /**
   * Format duration for display
   */
  public formatDuration(minutes: number): string {
    if (minutes < 60) {
      return `${minutes} min`;
    }
    const hours = Math.floor(minutes / 60);
    const remainingMinutes = minutes % 60;
    if (remainingMinutes === 0) {
      return `${hours} hr${hours > 1 ? 's' : ''}`;
    }
    return `${hours} hr${hours > 1 ? 's' : ''} ${remainingMinutes} min`;
  }

  /**
   * Get provider branding info
   */
  public getProviderBranding(providerName: string): { color: string; logo: string } | null {
    const provider = ExternalProviders.find(
      p => p.name.toLowerCase() === providerName.toLowerCase()
    );

    if (provider) {
      return {
        color: provider.primaryColor,
        logo: provider.logoUrl
      };
    }

    return null;
  }
}

export default ExternalTrainingService;
