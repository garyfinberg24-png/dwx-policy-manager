import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import { graphfi, GraphFI, SPFx as GraphSPFx } from '@pnp/graph';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/batching';

let _sp: SPFI | null = null;
let _graph: GraphFI | null = null;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context) {
    // Always reinitialize when context is provided to ensure fresh SP instance
    _sp = spfi().using(SPFx(context));
  }
  if (!_sp) {
    throw new Error('SP context not initialized. Please call getSP with a valid WebPartContext first.');
  }
  return _sp;
};

export const getGraph = (context?: WebPartContext): GraphFI => {
  if (context && !_graph) {
    _graph = graphfi().using(GraphSPFx(context));
  }
  return _graph!;
};
