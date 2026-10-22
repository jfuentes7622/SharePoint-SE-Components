export interface ICapacitySnapshot {
  capacity: number;
  claimed: number;
  waitlistCapacity: number;
  waitlisted: number;
}

export function availableSeats(snapshot: ICapacitySnapshot): number {
  return Math.max(0, Number(snapshot.capacity || 0) - Number(snapshot.claimed || 0));
}

export function availableWaitlistSeats(snapshot: ICapacitySnapshot): number {
  return Math.max(0, Number(snapshot.waitlistCapacity || 0) - Number(snapshot.waitlisted || 0));
}

export function canConfirm(snapshots: ICapacitySnapshot[], requestedSeats: number): boolean {
  return requestedSeats > 0 && snapshots.length > 0 && snapshots.every(function(snapshot: ICapacitySnapshot): boolean {
    return availableSeats(snapshot) >= requestedSeats;
  });
}

export function canWaitlist(snapshots: ICapacitySnapshot[], requestedSeats: number): boolean {
  return requestedSeats > 0 && snapshots.length > 0 && snapshots.every(function(snapshot: ICapacitySnapshot): boolean {
    return availableWaitlistSeats(snapshot) >= requestedSeats;
  });
}

export function validateCapacityDecrease(newCapacity: number, claimedCount: number): void {
  if (newCapacity < 0 || Math.floor(newCapacity) !== newCapacity) {
    throw new Error('Capacity must be a non-negative whole number.');
  }
  if (newCapacity < claimedCount) {
    throw new Error('Capacity cannot be lower than the ' + String(claimedCount) + ' claimed seats.');
  }
}
