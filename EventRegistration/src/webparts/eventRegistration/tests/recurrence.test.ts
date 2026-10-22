import { assert } from 'chai';
import { generateOccurrences, occurrenceKey } from '../helpers/recurrence';
import { availableSeats, canConfirm, canWaitlist, validateCapacityDecrease } from '../helpers/capacity';

describe('Event Registration helpers', function(): void {
  it('creates deterministic occurrence keys', function(): void {
    assert.equal(occurrenceKey(7, new Date('2026-01-02T15:30:00Z')), '7-20260102T153000Z');
  });

  it('generates interval daily occurrences', function(): void {
    var rows = generateOccurrences(3, new Date('2026-01-01T10:00:00Z'), new Date('2026-01-01T11:00:00Z'), {
      type: 'daily',
      interval: 2,
      weekdays: [],
      until: new Date('2026-01-05T10:00:00Z')
    });
    assert.deepEqual(rows.map(function(row): number { return row.start.getUTCDate(); }), [1, 3, 5]);
  });

  it('generates selected weekly weekdays without duplicates', function(): void {
    var rows = generateOccurrences(4, new Date('2026-01-05T14:00:00Z'), new Date('2026-01-05T15:00:00Z'), {
      type: 'weekly',
      interval: 1,
      weekdays: [1, 3, 3],
      until: new Date('2026-01-14T14:00:00Z')
    });
    assert.equal(rows.length, 4);
  });

  it('evaluates confirmed and waitlist capacity across sessions', function(): void {
    var rows = [
      { capacity: 5, claimed: 3, waitlistCapacity: 4, waitlisted: 0 },
      { capacity: 6, claimed: 4, waitlistCapacity: 4, waitlisted: 1 }
    ];
    assert.equal(availableSeats(rows[0]), 2);
    assert.isTrue(canConfirm(rows, 2));
    assert.isFalse(canConfirm(rows, 3));
    assert.isTrue(canWaitlist(rows, 3));
    assert.throws(function(): void { validateCapacityDecrease(2, 3); }, /claimed seats/);
  });
});
