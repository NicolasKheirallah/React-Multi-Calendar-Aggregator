export class ColorUtils {
  /**
   * Default color palette for calendars
   */
  private static readonly DEFAULT_COLORS = [
    '#0078d4', // Blue
    '#038387', // Teal
    '#00bcf2', // Light Blue
    '#40e0d0', // Turquoise
    '#008272', // Dark Teal
    '#107c10', // Green
    '#bad80a', // Lime
    '#ffb900', // Yellow
    '#ff8c00', // Orange
    '#d13438', // Red
    '#e3008c', // Pink
    '#881798', // Purple
    '#8764b8', // Light Purple
    '#00188f', // Dark Blue
    '#002050', // Navy
    '#5c2d91', // Dark Purple
    '#ca5010', // Dark Orange
    '#986f0b', // Dark Yellow
    '#498205', // Dark Green
    '#004b1c'  // Dark Forest
  ];

  /**
   * Microsoft 365 theme colors
   */
  private static readonly M365_COLORS = {
    primary: '#0078d4',
    primaryDark: '#106ebe',
    primaryLight: '#c7e0f4',
    success: '#107c10',
    warning: '#ff8c00',
    error: '#d13438',
    info: '#0078d4',
    neutral: '#323130',
    neutralLight: '#edebe9',
    neutralLighter: '#f3f2f1'
  };

  /**
   * Exchange calendar color mappings
   */
  private static readonly EXCHANGE_COLOR_MAP: { [key: string]: string } = {
    'lightBlue': '#0078d4',
    'lightGreen': '#107c10',
    'lightOrange': '#ff8c00',
    'lightGray': '#737373',
    'lightYellow': '#ffb900',
    'lightTeal': '#038387',
    'lightPink': '#e3008c',
    'lightBrown': '#8e562e',
    'lightRed': '#d13438',
    'maxColor': '#881798',
    'auto': '#0078d4'
  };

  /**
   * Generate a color based on string hash
   */
  public static generateColorFromString(str: string): string {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
      hash = str.charCodeAt(i) + ((hash << 5) - hash);
      hash = hash & hash; // Convert to 32bit integer
    }
    
    const index = Math.abs(hash) % this.DEFAULT_COLORS.length;
    return this.DEFAULT_COLORS[index];
  }

  /**
   * Get color for calendar by index
   */
  public static getColorByIndex(index: number): string {
    return this.DEFAULT_COLORS[index % this.DEFAULT_COLORS.length];
  }

  /**
   * Map Exchange calendar color to hex
   */
  public static mapExchangeColor(exchangeColor: string): string {
    return this.EXCHANGE_COLOR_MAP[exchangeColor] || this.M365_COLORS.primary;
  }

  /**
   * Get contrasting text color (black or white) for background
   */
  public static getContrastingTextColor(backgroundColor: string): string {
    const rgb = this.hexToRgb(backgroundColor);
    if (!rgb) return '#ffffff';
    
    // Calculate luminance
    const luminance = (0.299 * rgb.r + 0.587 * rgb.g + 0.114 * rgb.b) / 255;
    
    return luminance > 0.5 ? '#000000' : '#ffffff';
  }

  /**
   * Convert hex color to RGB
   */
  public static hexToRgb(hex: string): { r: number; g: number; b: number } | undefined {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
      r: parseInt(result[1], 16),
      g: parseInt(result[2], 16),
      b: parseInt(result[3], 16)
    } : undefined;
  }

  /**
   * Convert RGB to hex
   */
  public static rgbToHex(r: number, g: number, b: number): string {
    return `#${((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1)}`;
  }

  /**
   * Lighten a color by percentage
   */
  public static lightenColor(color: string, percent: number): string {
    const rgb = this.hexToRgb(color);
    if (!rgb) return color;
    
    const newR = Math.min(255, Math.floor(rgb.r + (255 - rgb.r) * (percent / 100)));
    const newG = Math.min(255, Math.floor(rgb.g + (255 - rgb.g) * (percent / 100)));
    const newB = Math.min(255, Math.floor(rgb.b + (255 - rgb.b) * (percent / 100)));
    
    return this.rgbToHex(newR, newG, newB);
  }

  /**
   * Darken a color by percentage
   */
  public static darkenColor(color: string, percent: number): string {
    const rgb = this.hexToRgb(color);
    if (!rgb) return color;
    
    const newR = Math.max(0, Math.floor(rgb.r * (1 - percent / 100)));
    const newG = Math.max(0, Math.floor(rgb.g * (1 - percent / 100)));
    const newB = Math.max(0, Math.floor(rgb.b * (1 - percent / 100)));
    
    return this.rgbToHex(newR, newG, newB);
  }

  /**
   * Get color with opacity
   */
  public static getColorWithOpacity(color: string, opacity: number): string {
    const rgb = this.hexToRgb(color);
    if (!rgb) return color;
    
    return `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${opacity})`;
  }

  /**
   * Validate if string is valid hex color
   */
  public static isValidHexColor(color: string): boolean {
    return /^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/.test(color);
  }

  /**
   * Get Microsoft 365 themed colors
   */
  public static getM365Colors(): typeof ColorUtils.M365_COLORS {
    return { ...this.M365_COLORS };
  }

  /**
   * Get default color palette
   */
  public static getDefaultColors(): string[] {
    return [...this.DEFAULT_COLORS];
  }

  /**
   * Generate gradient CSS for calendar events
   */
  public static generateEventGradient(baseColor: string): string {
    const lightColor = this.lightenColor(baseColor, 20);
    return `linear-gradient(135deg, ${baseColor} 0%, ${lightColor} 100%)`;
  }

  /**
   * Get semantic color for event importance
   */
  public static getImportanceColor(importance: string): string {
    switch (importance?.toLowerCase()) {
      case 'high':
        return this.M365_COLORS.error;
      case 'normal':
        return this.M365_COLORS.primary;
      case 'low':
        return this.M365_COLORS.neutralLight;
      default:
        return this.M365_COLORS.primary;
    }
  }

  /**
   * Get status color for event response
   */
  public static getResponseColor(response: string): string {
    switch (response?.toLowerCase()) {
      case 'accepted':
        return this.M365_COLORS.success;
      case 'declined':
        return this.M365_COLORS.error;
      case 'tentative':
        return this.M365_COLORS.warning;
      case 'none':
      default:
        return this.M365_COLORS.neutral;
    }
  }

  /**
   * Generate color scheme for calendar type
   */
  public static getCalendarTypeColors(type: 'SharePoint' | 'Exchange'): { primary: string; secondary: string; accent: string } {
    switch (type) {
      case 'SharePoint':
        return {
          primary: '#0078d4',
          secondary: '#c7e0f4',
          accent: '#106ebe'
        };
      case 'Exchange':
        return {
          primary: '#0078d4',
          secondary: '#e1f5fe',
          accent: '#005a9e'
        };
      default:
        return {
          primary: this.M365_COLORS.primary,
          secondary: this.M365_COLORS.primaryLight,
          accent: this.M365_COLORS.primaryDark
        };
    }
  }

  /**
   * Generate accessible color combinations
   */
  public static generateAccessiblePair(baseColor: string): { background: string; text: string } {
    const textColor = this.getContrastingTextColor(baseColor);
    return {
      background: baseColor,
      text: textColor
    };
  }

  /**
   * Get color for time-based events
   */
  public static getTimeBasedColor(date: Date): string {
    const now = new Date();
    const diffMs = date.getTime() - now.getTime();
    const diffDays = diffMs / (1000 * 60 * 60 * 24);
    
    if (diffDays < 0) {
      return this.M365_COLORS.neutralLight; // Past events
    } else if (diffDays < 1) {
      return this.M365_COLORS.warning; // Today
    } else if (diffDays < 7) {
      return this.M365_COLORS.info; // This week
    } else {
      return this.M365_COLORS.primary; // Future
    }
  }

  /**
   * Generate random color from palette
   */
  public static getRandomColor(): string {
    const randomIndex = Math.floor(Math.random() * this.DEFAULT_COLORS.length);
    return this.DEFAULT_COLORS[randomIndex];
  }

  /**
   * Color distance calculation for better contrast
   */
  public static calculateColorDistance(color1: string, color2: string): number {
    const rgb1 = this.hexToRgb(color1);
    const rgb2 = this.hexToRgb(color2);
    
    if (!rgb1 || !rgb2) return 0;
    
    const rDiff = rgb1.r - rgb2.r;
    const gDiff = rgb1.g - rgb2.g;
    const bDiff = rgb1.b - rgb2.b;
    
    return Math.sqrt(rDiff * rDiff + gDiff * gDiff + bDiff * bDiff);
  }

  /**
   * Get optimal colors for multiple calendars (ensuring good contrast)
   */
  public static getOptimalColorSet(count: number): string[] {
    if (count <= this.DEFAULT_COLORS.length) {
      return this.DEFAULT_COLORS.slice(0, count);
    }
    
    // For more than available colors, generate variations
    const colors: string[] = [];
    const baseColors = this.DEFAULT_COLORS;
    
    for (let i = 0; i < count; i++) {
      const baseIndex = i % baseColors.length;
      const variation = Math.floor(i / baseColors.length);
      
      let color = baseColors[baseIndex];
      
      // Apply variations
      if (variation > 0) {
        if (variation % 2 === 1) {
          color = this.lightenColor(color, 30);
        } else {
          color = this.darkenColor(color, 20);
        }
      }
      
      colors.push(color);
    }
    
    return colors;
  }
}