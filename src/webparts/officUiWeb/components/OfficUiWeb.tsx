import * as React from 'react';

import type { IOfficUiWebProps } from './IOfficUiWebProps';

import {
  DocumentCard,
  DocumentCardPreview,
  DocumentCardTitle,
  DocumentCardActivity,
  IDocumentCardPreviewProps,
  Icon 
} from 'office-ui-fabric-react/lib';

import { AnimationClassNames } from '@uifabric/styling';

export default class OfficUiWeb extends React.Component<IOfficUiWebProps, {}> {
  public render(): JSX.Element {
    const previewProps: IDocumentCardPreviewProps = {
      previewImages: [
        {
          previewImageSrc: String(require('./document-preview.png')),
          iconSrc: String(require('./icon-ppt.png')),
          width: 318,
          height: 196,
          accentColor: '#ce4b1f'
        }
      ],
    };
    const fadeInAnimationClassNames = AnimationClassNames.fadeIn500;
    const fadeInAnimationClassNames02 = AnimationClassNames.fadeOut500;
    //...previewPros is a spread operator that is used to pass all prop to previewProps
    return (
      <DocumentCard onClickHref='http://bing.com' className={fadeInAnimationClassNames}>
        <DocumentCardPreview { ...previewProps } />  
        <DocumentCardTitle title='Revenue stream proposal fiscal year 2016 version02.pptx' />
        
        <DocumentCardActivity
          activity='Created Feb 23, 2016'
          people={
            [
              { name: 'Kat Larrson', profileImageSrc: String(require('./avatar-kat.png')) }
            ]
          }
        />
        <Icon className={fadeInAnimationClassNames02} iconName="ExcelDocument" styles={{ root: { fontSize: 32 } }} />
        <Icon iconName="WordDocument" styles={{ root: { fontSize: 32 } }} />
        <Icon iconName="OneNoteLogoInverse" styles={{ root: { fontSize: 32 } }} />
        
      </DocumentCard>
    );
  }
}
