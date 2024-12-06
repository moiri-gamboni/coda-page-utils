import * as coda from "@codahq/packs-sdk";

export const pack = coda.newPack();
pack.addNetworkDomain("coda.io");

pack.setSystemAuthentication({
  type: coda.AuthenticationType.HeaderBearerToken,
});

const PAGE_SEARCH_FN = async (context, search, parameters) => {
  let pagesData = [];
  let continuationHref : string;

  do {
    let response = await context.fetcher.fetch({
      method: "GET",
      url: continuationHref || `https://coda.io/apis/v1/docs/${context.invocationLocation.docId}/pages?limit=100`,
      cacheTtlSecs: 0
    });
    continuationHref = response.body.nextPageLink;
    
    for (let page of response.body.items) {
      pagesData.push(page);
    }
  } while (!!continuationHref)

  return coda.autocompleteSearchObjects(search, pagesData, "name", "id");
};

const ICON_SEARCH_FN = async (context, search, parameters) => {
  let iconsData: coda.MetadataFormulaObjectResultType[] = [];

  let response = await context.fetcher.fetch({
    method: "GET",
    url: coda.withQueryParams('https://coda.io/api/icons', {
      term: search,
      limit: 50
    })
  });
  
  for (let icon of response.body.icons) {
    iconsData.push({
      value: icon.name,
      display: icon.label
    });
  }

  return iconsData;
};

pack.addFormula({
  name: "ListPages",
  description: "Returns a list of all pages in the doc as [ID, name] pairs",
  parameters: [
    coda.makeParameter({
      type: coda.ParameterType.Number,
      name: "limit",
      description: "Maximum number of pages to return (default: 100)",
      optional: true
    })
  ],
  resultType: coda.ValueType.Array,
  items: coda.makeSchema({
    type: coda.ValueType.Array,
    items: { type: coda.ValueType.String }
  }),
  execute: async function ([limit], context) {
    const response = await context.fetcher.fetch({
      method: "GET",
      url: coda.withQueryParams(`https://coda.io/apis/v1/docs/${context.invocationLocation.docId}/pages`, {
        limit: limit || 100
      }),
      cacheTtlSecs: 0
    });
    
    let pages = response.body.items.map(page => [page.id, page.name]);
    return pages;
  },
  cacheTtlSecs: 0
});

pack.addFormula({
  name: "AddPage",
  description: "Add a new page. Upon execution, returns an ID of the created page that you can then use to update it.",
  parameters: [
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "name",
      description: "Name of the page",
      optional: true
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "parent",
      description: "Parent of this new page",
      optional: true,
      autocomplete: PAGE_SEARCH_FN
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "subtitle",
      description: "Subtitle of the page",
      optional: true
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "iconName",
      description: "Name of the icon for this new page",
      optional: true,
      autocomplete: ICON_SEARCH_FN
    }),
    coda.makeParameter({
      type: coda.ParameterType.Image,
      name: "coverImage",
      description: "Cover image to use",
      optional: true
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "content",
      description: "Content of the page in Markdown format",
      optional: true
    }),
  ],
  resultType: coda.ValueType.String,
  isAction: true,
  execute: async function ([name, parentPageId, subtitle, iconName, imageUrl, content], context) {
    let payload: { 
      name: string; 
      parentPageId: string; 
      subtitle: string; 
      iconName: string; 
      imageUrl: string;
      pageContent?: {
        type: "canvas";
        canvasContent: {
          content: string;
          format: "markdown";
        };
      };
    } = { name, parentPageId, subtitle, iconName, imageUrl };
    if (content) {
      payload.pageContent = {
        type: "canvas",
        canvasContent: {
          content,
          format: "markdown"
        }
      };
    }
    Object.keys(payload).forEach((k) => {
      if (payload[k] === undefined) {
        delete payload[k];
      }
    });
    let result = await context.fetcher.fetch({
      method: 'POST',
      url: `https://coda.io/apis/v1/docs/${context.invocationLocation.docId}/pages`,
      body: JSON.stringify(payload),
      headers: {
        'Content-Type': 'application/json'
      }
    });
    return result.body.id
  },
});

pack.addFormula({
  name: "RenamePage",
  description: "Rename an existing page",
  parameters: [    
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "pageIdOrName",
      description: "ID or name of the page to rename. Prefer using IDs because names can change and there can be multiple pages with the same name. Use autocomplete to select the page, or use the result of the earlier AddPage() call, or switch on Developer Mode to get page IDs from the context menu",
      autocomplete: PAGE_SEARCH_FN
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "name",
      description: "New name of the page",
      optional: true
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "subtitle",
      description: "New subtitle of the page",
      optional: true
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "iconName",
      description: "New name of the icon for this page",
      optional: true,
      autocomplete: ICON_SEARCH_FN
    }),
    coda.makeParameter({
      type: coda.ParameterType.Image,
      name: "coverImage",
      description: "New cover image to use",
      optional: true
    }),
  ],
  resultType: coda.ValueType.String,
  isAction: true,
  execute: async function ([pageIdOrName, name, subtitle, iconName, imageUrl], context) {
    let payload = { name, subtitle, iconName, imageUrl };
    Object.keys(payload).forEach((k) => {
      if (payload[k] === undefined) {
        delete payload[k];
      }
    });
    let result = await context.fetcher.fetch({
      method: 'PUT',
      url: `${context.endpoint}/pages/${encodeURIComponent(pageIdOrName)}`,
      body: JSON.stringify(payload),
      headers: {
        'Content-Type': 'application/json'
      }
    });
    return result.body.id
  },
});
