/**
 * AudienceService Unit Tests — CRUD, evaluateAudience, evaluateAndSave
 */
import { AudienceService } from '../AudienceService';

// -- PnP SP mock builder --
function createMockSp(resolved: unknown = [], addResult?: unknown) {
  const getByIdMock = jest.fn().mockReturnValue({
    select: jest.fn().mockReturnValue(jest.fn().mockResolvedValue(resolved)),
    update: jest.fn().mockResolvedValue(undefined),
    delete: jest.fn().mockResolvedValue(undefined),
  });
  const itemsMock = {
    select: jest.fn().mockReturnThis(), filter: jest.fn().mockReturnThis(),
    orderBy: jest.fn().mockReturnThis(),
    top: jest.fn().mockReturnValue(jest.fn().mockResolvedValue(resolved)),
    add: jest.fn().mockResolvedValue(addResult ?? { data: resolved }),
    getById: getByIdMock,
  };
  const sp = { web: { lists: { getByTitle: jest.fn().mockReturnValue({ items: itemsMock }) } } };
  return { sp, itemsMock, getByIdMock };
}

function createFailSp() {
  const rej = jest.fn().mockRejectedValue(new Error('SP Error'));
  return { web: { lists: { getByTitle: jest.fn().mockReturnValue({ items: {
    select: jest.fn().mockReturnThis(), filter: jest.fn().mockReturnThis(),
    orderBy: jest.fn().mockReturnThis(), top: jest.fn().mockReturnValue(rej),
    getById: jest.fn().mockReturnValue({ select: jest.fn().mockReturnValue(rej) }),
  } }) } } };
}

describe('AudienceService', () => {
  let service: AudienceService;
  let m: ReturnType<typeof createMockSp>;
  beforeEach(() => { jest.clearAllMocks(); m = createMockSp(); service = new AudienceService(m.sp as any); });

  it('constructs without throwing', () => { expect(service).toBeDefined(); });

  describe('getAudiences', () => {
    it('queries PM_Audiences, maps nulls to defaults', async () => {
      const { sp } = createMockSp([
        { Id: 1, Title: 'A', Description: 'D', Criteria: '{"filters":[],"operator":"AND"}', MemberCount: 5, IsActive: true, LastEvaluated: '2026-01-01' },
        { Id: 2, Title: 'B', Description: null, Criteria: null, MemberCount: null, IsActive: false, LastEvaluated: null },
      ]);
      service = new AudienceService(sp as any);
      const r = await service.getAudiences();
      expect(sp.web.lists.getByTitle).toHaveBeenCalledWith('PM_Audiences');
      expect(r).toHaveLength(2);
      expect(r[1].Description).toBe('');
      expect(r[1].MemberCount).toBe(0);
    });
    it('returns [] on error', async () => {
      service = new AudienceService(createFailSp() as any);
      expect(await service.getAudiences()).toEqual([]);
    });
  });

  describe('getAudience', () => {
    it('calls getById and maps item', async () => {
      const item = { Id: 5, Title: 'HR', Description: '', Criteria: null, MemberCount: 10, IsActive: true, LastEvaluated: null };
      const { sp, getByIdMock } = createMockSp(item);
      getByIdMock.mockReturnValue({ select: jest.fn().mockReturnValue(jest.fn().mockResolvedValue(item)), update: jest.fn(), delete: jest.fn() });
      service = new AudienceService(sp as any);
      expect((await service.getAudience(5))?.Title).toBe('HR');
      expect(getByIdMock).toHaveBeenCalledWith(5);
    });
    it('returns null on error', async () => {
      service = new AudienceService(createFailSp() as any);
      expect(await service.getAudience(999)).toBeNull();
    });
  });

  describe('createAudience', () => {
    it('adds item with serialized Criteria, returns Id from data', async () => {
      const criteria = { filters: [{ field: 'Department' as const, operator: 'equals' as const, value: 'IT' }], operator: 'AND' as const };
      const { sp, itemsMock } = createMockSp([], { data: { Id: 42 } });
      service = new AudienceService(sp as any);
      const r = await service.createAudience({ Title: 'IT', Description: '', Criteria: criteria, MemberCount: 0, IsActive: true });
      expect(r.Id).toBe(42);
      expect(itemsMock.add.mock.calls[0][0].Criteria).toBe(JSON.stringify(criteria));
    });
  });

  describe('updateAudience', () => {
    it('calls getById().update with only defined fields', async () => {
      const upd = jest.fn().mockResolvedValue(undefined);
      m.getByIdMock.mockReturnValue({ update: upd, delete: jest.fn() });
      await service.updateAudience(3, { Title: 'New', IsActive: false });
      expect(m.getByIdMock).toHaveBeenCalledWith(3);
      expect(upd).toHaveBeenCalledWith({ Title: 'New', IsActive: false });
    });
  });

  describe('deleteAudience', () => {
    it('calls getById().delete', async () => {
      const del = jest.fn().mockResolvedValue(undefined);
      m.getByIdMock.mockReturnValue({ update: jest.fn(), delete: del });
      await service.deleteAudience(7);
      expect(m.getByIdMock).toHaveBeenCalledWith(7);
      expect(del).toHaveBeenCalled();
    });
  });

  describe('evaluateAudience', () => {
    it('queries PM_Employees, returns count + preview', async () => {
      const { sp } = createMockSp([{ Id: 1, Title: 'Alice', Email: 'a@co.com', Department: 'HR', JobTitle: 'Dev' }]);
      service = new AudienceService(sp as any);
      const r = await service.evaluateAudience({ filters: [], operator: 'AND' });
      expect(sp.web.lists.getByTitle).toHaveBeenCalledWith('PM_Employees');
      expect(r.count).toBe(1);
      expect(r.preview[0].Title).toBe('Alice');
    });
    it('returns empty result on error', async () => {
      service = new AudienceService(createFailSp() as any);
      expect(await service.evaluateAudience({ filters: [{ field: 'Department', operator: 'equals', value: 'X' }], operator: 'AND' }))
        .toEqual({ count: 0, preview: [] });
    });
  });

  describe('evaluateAndSave', () => {
    it('evaluates then persists MemberCount and LastEvaluated', async () => {
      const upd = jest.fn().mockResolvedValue(undefined);
      const gbi = jest.fn().mockReturnValue({ update: upd, delete: jest.fn() });
      const sp = { web: { lists: { getByTitle: jest.fn().mockReturnValue({ items: {
        select: jest.fn().mockReturnThis(), filter: jest.fn().mockReturnThis(),
        orderBy: jest.fn().mockReturnThis(),
        top: jest.fn().mockReturnValue(jest.fn().mockResolvedValue([{ Id: 1, Title: 'U', Email: 'u@co.com', Department: 'IT', JobTitle: 'D' }])),
        getById: gbi,
      } }) } } };
      service = new AudienceService(sp as any);
      const r = await service.evaluateAndSave(10, { filters: [], operator: 'AND' });
      expect(r.count).toBe(1);
      expect(gbi).toHaveBeenCalledWith(10);
      expect(upd).toHaveBeenCalledWith(expect.objectContaining({ MemberCount: 1 }));
    });
  });
});
