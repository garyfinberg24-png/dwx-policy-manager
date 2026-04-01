/**
 * EventClassifier — Matches raw event messages against EventCode patterns
 * to assign classification codes (e.g. APP-001, NET-010, SEC-001).
 *
 * Console events are pre-classified by ConsoleInterceptor (CON-001..005).
 * Network events are pre-classified by NetworkInterceptor (NET-xxx).
 * This classifier runs on Application/System/Audit/DLQ channel events
 * to assign more specific codes based on message content.
 */

import {
  IEventEntry,
  IEventClassificationResult,
  EventChannel,
} from '../../models/IEventViewer';
import { EVENT_CODES } from '../../constants/EventCodes';

export class EventClassifier {
  /**
   * Classify an event by matching its message against known patterns.
   * Returns a classification result, or undefined if no pattern matches.
   *
   * Note: Console and Network events already have codes assigned by their
   * respective interceptors. This method is primarily for Application,
   * System, Audit, and DLQ channel events.
   */
  public static classify(event: IEventEntry): IEventClassificationResult | undefined {
    // Skip if event already has a specific code (not a generic console/network code)
    if (event.eventCode && event.channel === EventChannel.Console) return undefined;
    if (event.eventCode && event.channel === EventChannel.Network) return undefined;

    const message = event.message || '';
    const stackTrace = event.stackTrace || '';
    const combined = `${message} ${stackTrace}`;

    // Skip generic console codes (CON-001..004) — they match everything
    for (let i = 0; i < EVENT_CODES.length; i++) {
      const codeDef = EVENT_CODES[i];

      // Skip console catch-all codes
      if (codeDef.code.startsWith('CON-')) continue;

      // Skip network codes — NetworkInterceptor handles these
      if (codeDef.code.startsWith('NET-')) continue;

      // Check each pattern
      for (let p = 0; p < codeDef.patterns.length; p++) {
        if (codeDef.patterns[p].test(combined)) {
          return {
            eventCode: codeDef.code,
            category: codeDef.category,
            description: codeDef.description,
            suggestedAction: codeDef.suggestedAction,
          };
        }
      }
    }

    return undefined;
  }

  /**
   * Re-classify a console event by checking if its message matches
   * a more specific APP/SEC/SYS code. If so, upgrades from the generic
   * CON-xxx code.
   */
  public static reclassifyConsoleEvent(event: IEventEntry): IEventClassificationResult | undefined {
    if (event.channel !== EventChannel.Console) return undefined;

    const message = event.message || '';
    const stackTrace = event.stackTrace || '';
    const combined = `${message} ${stackTrace}`;

    for (let i = 0; i < EVENT_CODES.length; i++) {
      const codeDef = EVENT_CODES[i];

      // Only try APP, SEC, SYS, DLQ codes
      if (codeDef.code.startsWith('CON-') || codeDef.code.startsWith('NET-')) continue;

      for (let p = 0; p < codeDef.patterns.length; p++) {
        if (codeDef.patterns[p].test(combined)) {
          return {
            eventCode: codeDef.code,
            category: codeDef.category,
            description: codeDef.description,
            suggestedAction: codeDef.suggestedAction,
          };
        }
      }
    }

    return undefined;
  }
}
