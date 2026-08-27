export default class PnPTelemetryNoop {
  private static instance: PnPTelemetryNoop = new PnPTelemetryNoop();

  public static getInstance(): PnPTelemetryNoop {
    return PnPTelemetryNoop.instance;
  }

  public optOut(): void {
    return;
  }

  public trackEvent(name: string, properties?: any): void {
    return;
  }
}