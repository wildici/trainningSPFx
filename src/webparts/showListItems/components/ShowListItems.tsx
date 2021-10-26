import * as React from 'react';
import styles from './ShowListItems.module.scss';
import { IShowListItemsProps } from './IShowListItemsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Form from './Form';

// export default class ShowListItems extends React.Component<IShowListItemsProps, {}> {
//   public render(): React.ReactElement<IShowListItemsProps> {
//     return (
//       <div className={ styles.showListItems }>
//         <div className={ styles.container }>
//           <div className={ styles.row }>
//             <div className={ styles.column }>
//               <span className={ styles.title }>Welcome to SharePoint!</span>
//               <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
//               <p className={ styles.description }>{escape(this.props.description)}</p>
//               <a href="https://aka.ms/spfx" className={ styles.button }>
//                 <span className={ styles.label }>Learn more</span>
//               </a>
//             </div>
//           </div>
//         </div>
//       </div>
//     );
//   }
// }

function ShowListItems() {
  return (
    <Form />
  )
}

export default ShowListItems